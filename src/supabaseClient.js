import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !key) {
  console.error(
    "⚠️ Variables Supabase manquantes. Créez un fichier .env (voir .env.example) " +
    "avec VITE_SUPABASE_URL et VITE_SUPABASE_ANON_KEY."
  );
}

export const supabase = createClient(url || "http://localhost", key || "anon", {
  auth: { persistSession: true, autoRefreshToken: true },
});

const TABLE = "commandes";

const FIELD_MAP = {
  date: "date", nom: "nom", prenom: "prenom", tel: "tel", email: "email",
  marque: "marque", modele: "modele", refInt: "ref_int", refFourn: "ref_fourn",
  couleur: "couleur", taille: "taille", etat: "etat", dateValid: "date_valid",
  vendeur: "vendeur", commentaire: "commentaire",
  articleCote: "article_cote", deCoteJusquau: "de_cote_jusquau",
  magasin: "magasin", archivedAt: "archived_at",
  msgVocal: "msg_vocal", msgWhatsapp: "msg_whatsapp", mailEnvoye: "mail_envoye",
};

function fromRow(r) {
  return {
    id: r.id,
    date: r.date || "", nom: r.nom || "", prenom: r.prenom || "",
    tel: r.tel || "", email: r.email || "",
    marque: r.marque || "", modele: r.modele || "",
    refInt: r.ref_int || "", refFourn: r.ref_fourn || "",
    couleur: r.couleur || "", taille: r.taille || "",
    etat: r.etat || "", dateValid: r.date_valid || "",
    vendeur: r.vendeur || "", commentaire: r.commentaire || "",
    articleCote: r.article_cote || "", deCoteJusquau: r.de_cote_jusquau || "",
    magasin: r.magasin || "SQUARE",
    archivedAt: r.archived_at || null,
    msgVocal: !!r.msg_vocal,
    msgWhatsapp: !!r.msg_whatsapp,
    mailEnvoye: !!r.mail_envoye,
  };
}

function toRow(partial) {
  const out = {};
  for (const k in partial) {
    if (FIELD_MAP[k] === undefined) continue;
    const v = partial[k];
    out[FIELD_MAP[k]] = v === undefined ? "" : v;
  }
  return out;
}

// ── AUTH ───────────────────────────────────────────────────────────────────
export async function signIn(email, password) {
  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) throw error;
  return data;
}
export async function signOut() { await supabase.auth.signOut(); }
export async function getSession() {
  const { data } = await supabase.auth.getSession();
  return data.session;
}
export function onAuthChange(callback) {
  return supabase.auth.onAuthStateChange((_event, session) => callback(session));
}
export async function resetPassword(email) {
  const { error } = await supabase.auth.resetPasswordForEmail(email, {
    redirectTo: window.location.origin,
  });
  if (error) throw error;
}
export async function updatePassword(newPassword) {
  const { error } = await supabase.auth.updateUser({ password: newPassword });
  if (error) throw error;
}

// ── PROFIL ────────────────────────────────────────────────────────────────
export async function fetchProfile(userId) {
  const { data, error } = await supabase
    .from("profiles").select("*").eq("id", userId).single();
  if (error) throw error;
  return data;
}

// ── COMMANDES ─────────────────────────────────────────────────────────────
export async function fetchOrders() {
  // RLS filtre automatiquement selon les magasins du profil
  const { data, error } = await supabase
    .from(TABLE).select("*").order("date", { ascending: false });
  if (error) throw error;
  return (data || []).map(fromRow);
}
export async function insertOrder(order) {
  const { data, error } = await supabase
    .from(TABLE).insert(toRow(order)).select().single();
  if (error) throw error;
  return fromRow(data);
}
export async function updateOrder(id, patch) {
  const { error } = await supabase
    .from(TABLE).update(toRow(patch)).eq("id", id);
  if (error) throw error;
}
export async function deleteOrder(id) {
  const { error } = await supabase.from(TABLE).delete().eq("id", id);
  if (error) throw error;
}
export async function deleteOrders(ids) {
  if (!ids || !ids.length) return;
  const { error } = await supabase.from(TABLE).delete().in("id", ids);
  if (error) throw error;
}
export async function bulkUpdateStatus(ids, etat, dateValid) {
  if (!ids || !ids.length) return;
  const { error } = await supabase
    .from(TABLE).update({ etat, date_valid: dateValid }).in("id", ids);
  if (error) throw error;
}
// Archiver (ne supprime pas, juste marque comme archivée)
export async function archiveOrders(ids) {
  if (!ids || !ids.length) return;
  const now = new Date().toISOString();
  const { error } = await supabase
    .from(TABLE).update({ archived_at: now }).in("id", ids);
  if (error) throw error;
  return now;
}
