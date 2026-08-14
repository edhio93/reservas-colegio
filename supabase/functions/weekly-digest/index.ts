import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

function json(data: unknown, status = 200): Response {
  return new Response(
    JSON.stringify(data),
    {
      status,
      headers: { "Content-Type": "application/json; charset=utf-8" },
    },
  );
}

function escapeHtml(value: unknown): string {
  return String(value ?? "—")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function relation(value: any): any {
  return Array.isArray(value) ? value[0] : value;
}

function chileNow() {
  const formatter = new Intl.DateTimeFormat(
    "en-US",
    {
      timeZone: "America/Santiago",
      weekday: "short",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hourCycle: "h23",
    },
  );

  const parts = Object.fromEntries(
    formatter
      .formatToParts(new Date())
      .filter((part) => part.type !== "literal")
      .map((part) => [part.type, part.value]),
  );

  return {
    weekday: parts.weekday,
    date: `${parts.year}-${parts.month}-${parts.day}`,
    hour: Number(parts.hour),
    minute: Number(parts.minute),
  };
}

function addDays(dateIso: string, days: number): string {
  const date = new Date(`${dateIso}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function formatDate(dateIso: string): string {
  const [year, month, day] = dateIso.split("-");
  return `${day}/${month}/${year}`;
}

function formatTime(value: unknown): string {
  return String(value || "").slice(0, 5);
}

Deno.serve(async (req) => {
  const current = chileNow();

  let force = false;
  let requestedWeekStart = "";

  try {
    const body = await req.json();
    force = Boolean(body?.force);
    requestedWeekStart = String(body?.week_start || "");
  } catch {
    // body opcional
  }

  // El cron corre cada 15 minutos.
  // Solo genera resúmenes los lunes entre 07:30 y 07:44 de Chile.
  // `force:true` existe exclusivamente para una prueba manual autenticada.
  if (
    !force &&
    (
      current.weekday !== "Mon" ||
      current.hour !== 7 ||
      current.minute < 30 ||
      current.minute > 44
    )
  ) {
    return json({
      ok: true,
      skipped: true,
      reason: "outside_local_window",
      chile: current,
    });
  }

  const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
  const serviceRole = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";

  const supabase = createClient(supabaseUrl, serviceRole, {
    auth: { persistSession: false },
  });

  const weekdayOffset: Record<string, number> = {
    Mon: 0,
    Tue: 1,
    Wed: 2,
    Thu: 3,
    Fri: 4,
    Sat: 5,
    Sun: 6,
  };

  const weekStart = requestedWeekStart || (
    current.weekday === "Mon"
      ? current.date
      : addDays(
          current.date,
          -(weekdayOffset[current.weekday] ?? 0),
        )
  );
  const weekEnd = addDays(weekStart, 4);

  const { data: reservations, error } = await supabase
    .from("reservas")
    .select(
      "id,fecha,hora_inicio,hora_fin,profesor,"
        + "profesores(id,nombre,email,workspace_primary_email,workspace_active),"
        + "cursos(nombre),recursos(nombre)",
    )
    .gte("fecha", weekStart)
    .lte("fecha", weekEnd)
    .order("fecha", { ascending: true })
    .order("hora_inicio", { ascending: true });

  if (error) {
    return json({ error: error.message }, 500);
  }

  const byProfessor = new Map<number, any[]>();

  for (const reservation of reservations || []) {
    const professor = relation(reservation.profesores);

    if (!professor?.id) continue;

    if (!byProfessor.has(professor.id)) {
      byProfessor.set(professor.id, []);
    }

    byProfessor.get(professor.id)?.push({
      ...reservation,
      professor,
      course: relation(reservation.cursos),
      resource: relation(reservation.recursos),
    });
  }

  let queued = 0;
  let skipped = 0;

  for (const [professorId, rows] of byProfessor.entries()) {
    // Esta condición garantiza: si no tiene reservas, no se envía nada.
    if (!rows.length) {
      skipped += 1;
      continue;
    }

    const professor = rows[0].professor;

    const recipient = String(
      professor.workspace_primary_email ||
      professor.email ||
      "",
    ).trim().toLowerCase();

    if (!recipient) {
      skipped += 1;
      continue;
    }

    const { data: existing } = await supabase
      .from("weekly_digest_log")
      .select("id,status")
      .eq("professor_id", professorId)
      .eq("week_start", weekStart)
      .maybeSingle();

    if (existing) {
      skipped += 1;
      continue;
    }

    const cards = rows.map((row) => `
      <tr>
        <td style="padding:10px;border-bottom:1px solid #e5e7eb">
          <b>${escapeHtml(formatDate(row.fecha))}</b><br>
          ${escapeHtml(formatTime(row.hora_inicio))}
          –${escapeHtml(formatTime(row.hora_fin))}
        </td>
        <td style="padding:10px;border-bottom:1px solid #e5e7eb">
          ${escapeHtml(row.course?.nombre)}
        </td>
        <td style="padding:10px;border-bottom:1px solid #e5e7eb">
          ${escapeHtml(row.resource?.nombre)}
        </td>
      </tr>
    `).join("");

    const subject =
      `📅 Tu semana en Enlaces · ${formatDate(weekStart)} al ${formatDate(weekEnd)}`;

    const body = `
    <html>
    <body style="font-family:Arial,sans-serif;color:#172033;line-height:1.55">
      <h2 style="color:#800020">Tu calendario semanal de Enlaces</h2>
      <p>Hola ${escapeHtml(professor.nombre)},</p>
      <p>
        Este es tu calendario de Enlaces para la semana del
        <b>${escapeHtml(formatDate(weekStart))}</b> al
        <b>${escapeHtml(formatDate(weekEnd))}</b>.
      </p>

      <table style="border-collapse:collapse;width:100%">
        <thead>
          <tr style="background:#f8fafc">
            <th style="text-align:left;padding:10px">Fecha / horario</th>
            <th style="text-align:left;padding:10px">Curso</th>
            <th style="text-align:left;padding:10px">Recurso</th>
          </tr>
        </thead>
        <tbody>${cards}</tbody>
      </table>

      <p><b>${rows.length}</b> reserva${rows.length === 1 ? "" : "s"} esta semana.</p>

      <p style="margin-top:24px">
        Departamento de Informática / Enlaces<br>
        Liceo Bicentenario de Excelencia Colegio Antonio Varas
      </p>
    </body>
    </html>`;

    const { error: logError } = await supabase
      .from("weekly_digest_log")
      .insert({
        professor_id: professorId,
        week_start: weekStart,
        recipient_email: recipient,
        reservations_count: rows.length,
        status: "queued",
      });

    if (logError) {
      // Dedupe/concurrencia: no duplicar.
      skipped += 1;
      continue;
    }

    const { error: outboxError } = await supabase
      .from("notification_outbox")
      .insert({
        type: "weekly_digest",
        professor_id: professorId,
        reservation_id: null,
        recipient_email: recipient,
        subject,
        html_body: body,
        metadata: {
          week_start: weekStart,
          week_end: weekEnd,
          reservations_count: rows.length,
        },
        dedupe_key: `weekly:${professorId}:${weekStart}`,
        status: "pending",
        attempts: 0,
        available_at: new Date().toISOString(),
      });

    if (outboxError) {
      await supabase
        .from("weekly_digest_log")
        .update({
          status: "error",
          error: outboxError.message,
        })
        .eq("professor_id", professorId)
        .eq("week_start", weekStart);

      skipped += 1;
      continue;
    }

    queued += 1;
  }

  return json({
    ok: true,
    week_start: weekStart,
    week_end: weekEnd,
    professors_with_reservations: byProfessor.size,
    queued,
    skipped,
  });
});
