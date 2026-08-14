import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { sendGmailHtml } from "../_shared/gmail.ts";

function json(data: unknown, status = 200): Response {
  return new Response(
    JSON.stringify(data),
    {
      status,
      headers: { "Content-Type": "application/json; charset=utf-8" },
    },
  );
}

function retryAt(attempts: number): string {
  const seconds = Math.min(
    3600,
    Math.max(60, Math.pow(2, attempts) * 60),
  );

  return new Date(Date.now() + seconds * 1000).toISOString();
}

Deno.serve(async () => {
  const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
  const serviceRole = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";

  if (!supabaseUrl || !serviceRole) {
    return json({ error: "Supabase environment missing" }, 500);
  }

  const supabase = createClient(supabaseUrl, serviceRole, {
    auth: { persistSession: false },
  });

  const now = new Date().toISOString();

  const { data: jobs, error } = await supabase
    .from("notification_outbox")
    .select(
      "id,type,professor_id,reservation_id,recipient_email,subject,"
        + "html_body,status,attempts,metadata",
    )
    .eq("status", "pending")
    .lte("available_at", now)
    .order("created_at", { ascending: true })
    .limit(20);

  if (error) {
    return json({ error: error.message }, 500);
  }

  let sent = 0;
  let failed = 0;

  for (const job of jobs || []) {
    const { data: locked, error: lockError } = await supabase
      .from("notification_outbox")
      .update({
        status: "sending",
        updated_at: new Date().toISOString(),
      })
      .eq("id", job.id)
      .eq("status", "pending")
      .select("id")
      .maybeSingle();

    if (lockError || !locked) continue;

    try {
      await sendGmailHtml({
        to: job.recipient_email,
        subject: job.subject,
        html: job.html_body,
      });

      const sentAt = new Date().toISOString();

      await supabase
        .from("notification_outbox")
        .update({
          status: "sent",
          attempts: Number(job.attempts || 0) + 1,
          error: null,
          sent_at: sentAt,
          updated_at: sentAt,
        })
        .eq("id", job.id);

      if (job.type === "weekly_digest" && job.professor_id) {
        const weekStart = job.metadata?.week_start;

        if (weekStart) {
          await supabase
            .from("weekly_digest_log")
            .update({
              status: "sent",
              sent_at: sentAt,
              error: null,
            })
            .eq("professor_id", job.professor_id)
            .eq("week_start", weekStart);
        }
      }

      sent += 1;
    } catch (sendError) {
      const attempts = Number(job.attempts || 0) + 1;
      const finalError =
        sendError instanceof Error
          ? sendError.message
          : String(sendError);

      const terminal = attempts >= 5;

      await supabase
        .from("notification_outbox")
        .update({
          status: terminal ? "error" : "pending",
          attempts,
          error: finalError,
          available_at: retryAt(attempts),
          updated_at: new Date().toISOString(),
        })
        .eq("id", job.id);

      if (
        terminal &&
        job.type === "weekly_digest" &&
        job.professor_id &&
        job.metadata?.week_start
      ) {
        await supabase
          .from("weekly_digest_log")
          .update({
            status: "error",
            error: finalError,
          })
          .eq("professor_id", job.professor_id)
          .eq("week_start", job.metadata.week_start);
      }

      failed += 1;
    }
  }

  return json({
    ok: true,
    considered: (jobs || []).length,
    sent,
    failed,
  });
});
