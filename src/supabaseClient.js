import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !key) {
  console.error(
    "⚠️ Variables Supabase manquantes. Créez un fichier .env (voir .env.example) " +
    "avec VITE_SUPABASE_URL et VITE_SUPABASE_ANON_KEY."
  );
}

export const supabase = createClient(url || "http://localhost", key || "anon");

const TABLE = "commandes";

// Correspondance champs appli (camelCase) ↔ colonnes base (snake_case)
const FIELD_MAP = {
  date: "date",
  nom: "nom",
  prenom: "prenom",
  tel: "tel",
  email: "email",
  marque: "marque",
  modele: "modele",
  refInt: "ref_int",
  refFourn: "ref_fourn",
  couleur: "couleur",
  taille: "taille",
  etat: "etat",
  dateValid: "date_valid",
  vendeur: "vendeur",
  commentaire: "commentaire",
  articleCote: "article_cote",
  deCoteJusquau: "de_cote_jusquau",
  dateRecup: "date_recup",
};

// base → appli
function fromRow(r) {
  return {
    id: r.id,
    date: r.date || "",
    nom: r.nom || "",
    prenom: r.prenom || "",
    tel: r.tel || "",
    email: r.email || "",
    marque: r.marque || "",
    modele: r.modele || "",
    refInt: r.ref_int || "",
    refFourn: r.ref_fourn || "",
    couleur: r.couleur || "",
    taille: r.taille || "",
    etat: r.etat || "",
    dateValid: r.date_valid || "",
    vendeur: r.vendeur || "",
    commentaire: r.commentaire || "",
    articleCote: r.article_cote || "",
    deCoteJusquau: r.de_cote_jusquau || "",
    dateRecup: r.date_recup || "",
  };
}

// appli (complet ou partiel) → base
function toRow(partial) {
  const out = {};
  for (const k in partial) {
    if (FIELD_MAP[k]) out[FIELD_MAP[k]] = partial[k] ?? "";
  }
  return out;
}

// ── API ────────────────────────────────────────────────────────────────────

export async function fetchOrders() {
  const { data, error } = await supabase
    .from(TABLE)
    .select("*")
    .order("date", { ascending: false });
  if (error) throw error;
  return (data || []).map(fromRow);
}

export async function insertOrder(order) {
  const { data, error } = await supabase
    .from(TABLE)
    .insert(toRow(order))
    .select()
    .single();
  if (error) throw error;
  return fromRow(data);
}

export async function updateOrder(id, patch) {
  const { error } = await supabase
    .from(TABLE)
    .update(toRow(patch))
    .eq("id", id);
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
    .from(TABLE)
    .update({ etat, date_valid: dateValid })
    .in("id", ids);
  if (error) throw error;
}
