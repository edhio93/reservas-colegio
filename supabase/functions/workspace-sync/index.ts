import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { getGoogleAccessToken } from "../_shared/google_auth.ts";

const DIRECTORY_SCOPES = [
  "https://www.googleapis.com/auth/admin.directory.user.readonly",
  "https://www.googleapis.com/auth/admin.directory.group.readonly",
  "https://www.googleapis.com/auth/admin.directory.group.member.readonly",
];

function json(data: unknown, status = 200): Response {
  return new Response(
    JSON.stringify(data),
    {
      status,
      headers: { "Content-Type": "application/json; charset=utf-8" },
    },
  );
}

function normalizeName(value: string): string {
  return String(value || "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ");
}

async function googleGet(
  url: URL,
  token: string,
): Promise<Record<string, unknown>> {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Google Directory ${response.status}: ${JSON.stringify(data)}`,
    );
  }

  return data;
}

async function listUsers(
  token: string,
  customer: string,
): Promise<any[]> {
  const rows: any[] = [];
  let pageToken = "";

  do {
    const url = new URL(
      "https://admin.googleapis.com/admin/directory/v1/users",
    );

    url.searchParams.set("customer", customer);
    url.searchParams.set("maxResults", "500");
    url.searchParams.set("orderBy", "email");
    url.searchParams.set("projection", "full");

    if (pageToken) url.searchParams.set("pageToken", pageToken);

    const data: any = await googleGet(url, token);
    rows.push(...(data.users || []));
    pageToken = data.nextPageToken || "";
  } while (pageToken);

  return rows;
}

async function listGroups(
  token: string,
  customer: string,
): Promise<any[]> {
  const rows: any[] = [];
  let pageToken = "";

  do {
    const url = new URL(
      "https://admin.googleapis.com/admin/directory/v1/groups",
    );

    url.searchParams.set("customer", customer);
    url.searchParams.set("maxResults", "200");
    url.searchParams.set("orderBy", "email");

    if (pageToken) url.searchParams.set("pageToken", pageToken);

    const data: any = await googleGet(url, token);
    rows.push(...(data.groups || []));
    pageToken = data.nextPageToken || "";
  } while (pageToken);

  return rows;
}

async function listMembers(
  token: string,
  groupKey: string,
): Promise<any[]> {
  const rows: any[] = [];
  let pageToken = "";

  do {
    const url = new URL(
      `https://admin.googleapis.com/admin/directory/v1/groups/${
        encodeURIComponent(groupKey)
      }/members`,
    );

    url.searchParams.set("maxResults", "200");

    if (pageToken) url.searchParams.set("pageToken", pageToken);

    const data: any = await googleGet(url, token);
    rows.push(...(data.members || []));
    pageToken = data.nextPageToken || "";
  } while (pageToken);

  return rows;
}

Deno.serve(async (req) => {
  const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
  const serviceRole = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";
  const delegatedAdmin = Deno.env.get("GOOGLE_DELEGATED_ADMIN") || "";
  const customer = Deno.env.get("GOOGLE_WORKSPACE_CUSTOMER") || "my_customer";

  if (!supabaseUrl || !serviceRole) {
    return json({ error: "Supabase environment missing" }, 500);
  }

  if (!delegatedAdmin) {
    return json(
      { error: "Falta GOOGLE_DELEGATED_ADMIN en Edge Function Secrets." },
      500,
    );
  }

  const supabase = createClient(supabaseUrl, serviceRole, {
    auth: { persistSession: false },
  });

  let source = "scheduled";

  try {
    const body = await req.json();
    source = body?.source || source;
  } catch {
    // body opcional
  }

  const startedAt = new Date().toISOString();

  const { data: logRow, error: logError } = await supabase
    .from("workspace_sync_log")
    .insert({
      source,
      status: "running",
      started_at: startedAt,
    })
    .select("id")
    .single();

  if (logError || !logRow) {
    return json({ error: `No se pudo crear sync log: ${logError?.message}` }, 500);
  }

  const syncId = logRow.id;

  try {
    const token = await getGoogleAccessToken(
      delegatedAdmin,
      DIRECTORY_SCOPES,
    );

    const users = await listUsers(token, customer);
    const groups = await listGroups(token, customer);

    const now = new Date().toISOString();

    if (users.length) {
      const rows = users.map((u: any) => ({
        google_id: u.id,
        primary_email: String(u.primaryEmail || "").toLowerCase(),
        full_name:
          u.name?.fullName ||
          [u.name?.givenName, u.name?.familyName].filter(Boolean).join(" "),
        given_name: u.name?.givenName || null,
        family_name: u.name?.familyName || null,
        org_unit_path: u.orgUnitPath || null,
        suspended: Boolean(u.suspended),
        archived: Boolean(u.archived),
        is_admin: Boolean(u.isAdmin),
        present_in_directory: true,
        raw: u,
        synced_at: now,
      }));

      const { error } = await supabase
        .from("workspace_users")
        .upsert(rows, { onConflict: "google_id" });

      if (error) throw error;
    }

    if (groups.length) {
      const rows = groups.map((g: any) => ({
        google_id: g.id,
        email: String(g.email || "").toLowerCase(),
        name: g.name || g.email || "Grupo",
        description: g.description || null,
        direct_members_count:
          Number.parseInt(String(g.directMembersCount || "0")) || 0,
        present_in_directory: true,
        raw: g,
        synced_at: now,
      }));

      const { error } = await supabase
        .from("workspace_groups")
        .upsert(rows, { onConflict: "google_id" });

      if (error) throw error;
    }

    // Marcar registros que no aparecieron en esta sincronización.
    await supabase
      .from("workspace_users")
      .update({ present_in_directory: false })
      .lt("synced_at", startedAt);

    await supabase
      .from("workspace_groups")
      .update({ present_in_directory: false })
      .lt("synced_at", startedAt);

    let membersCount = 0;

    for (const group of groups) {
      const members = await listMembers(token, group.id);

      const { error: deleteError } = await supabase
        .from("workspace_group_members")
        .delete()
        .eq("group_google_id", group.id);

      if (deleteError) throw deleteError;

      if (members.length) {
        const rows = members.map((member: any) => ({
          group_google_id: group.id,
          member_google_id: member.id || null,
          member_email: String(member.email || "").toLowerCase(),
          role: member.role || null,
          type: member.type || null,
          status: member.status || null,
          synced_at: now,
        }));

        const { error } = await supabase
          .from("workspace_group_members")
          .upsert(rows, {
            onConflict: "group_google_id,member_email",
          });

        if (error) throw error;
        membersCount += rows.length;
      }
    }

    // ------------------------------------------------------------
    // Vinculación automática profesor -> usuario Workspace
    // Orden:
    // 1) workspace_user_id existente
    // 2) email actual exacto
    // 3) nombre completo exacto, solo si hay una única coincidencia
    // ------------------------------------------------------------
    const { data: professors, error: professorsError } = await supabase
      .from("profesores")
      .select(
        "id,nombre,email,workspace_user_id,workspace_primary_email",
      );

    if (professorsError) throw professorsError;

    const userById = new Map<string, any>();
    const userByEmail = new Map<string, any>();
    const usersByName = new Map<string, any[]>();

    for (const user of users) {
      userById.set(String(user.id), user);
      userByEmail.set(
        String(user.primaryEmail || "").toLowerCase(),
        user,
      );

      const normalized = normalizeName(
        user.name?.fullName ||
        [user.name?.givenName, user.name?.familyName]
          .filter(Boolean)
          .join(" "),
      );

      if (!usersByName.has(normalized)) {
        usersByName.set(normalized, []);
      }
      usersByName.get(normalized)?.push(user);
    }

    let linked = 0;

    for (const professor of professors || []) {
      let user: any | undefined;
      let method: string | undefined;

      if (
        professor.workspace_user_id &&
        userById.has(String(professor.workspace_user_id))
      ) {
        user = userById.get(String(professor.workspace_user_id));
        method = "google_id";
      }

      if (!user && professor.email) {
        user = userByEmail.get(
          String(professor.email).toLowerCase().trim(),
        );
        if (user) method = "email";
      }

      if (!user) {
        const candidates =
          usersByName.get(normalizeName(professor.nombre)) || [];

        if (candidates.length === 1) {
          user = candidates[0];
          method = "name_unique";
        }
      }

      if (!user) continue;

      const email = String(user.primaryEmail || "").toLowerCase();

      const { error } = await supabase
        .from("profesores")
        .update({
          email,
          workspace_user_id: user.id,
          workspace_primary_email: email,
          workspace_active: !Boolean(user.suspended),
          workspace_org_unit: user.orgUnitPath || null,
          workspace_match_method: method,
          workspace_last_sync: now,
        })
        .eq("id", professor.id);

      if (error) throw error;
      linked += 1;
    }

    await supabase
      .from("workspace_sync_log")
      .update({
        status: "success",
        users_count: users.length,
        groups_count: groups.length,
        members_count: membersCount,
        linked_professors_count: linked,
        finished_at: new Date().toISOString(),
      })
      .eq("id", syncId);

    return json({
      ok: true,
      users: users.length,
      groups: groups.length,
      members: membersCount,
      linked_professors: linked,
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : String(error);

    await supabase
      .from("workspace_sync_log")
      .update({
        status: "error",
        error: message,
        finished_at: new Date().toISOString(),
      })
      .eq("id", syncId);

    return json({ ok: false, error: message }, 500);
  }
});
