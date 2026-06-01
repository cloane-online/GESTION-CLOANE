-- ════════════════════════════════════════════════════════════════════════════
--  CLOANE — Schéma de la base Supabase
--  À exécuter EN PREMIER dans  Supabase → SQL Editor → New query → Run
-- ════════════════════════════════════════════════════════════════════════════

-- Table des commandes clients B2B
create table if not exists public.commandes (
  id               uuid primary key default gen_random_uuid(),
  created_at       timestamptz not null default now(),
  updated_at       timestamptz not null default now(),

  date             text default '',   -- date de création (ISO yyyy-mm-dd)
  nom              text default '',
  prenom           text default '',
  tel              text default '',
  email            text default '',
  marque           text default '',
  modele           text default '',
  ref_int          text default '',
  ref_fourn        text default '',
  couleur          text default '',
  taille           text default '',
  etat             text default '',
  date_valid       text default '',   -- date de validation (ISO)
  vendeur          text default '',
  commentaire      text default '',
  article_cote     text default '',   -- article réservé (client prévenu)
  de_cote_jusquau  text default '',   -- date limite de mise de côté (ISO)
  date_recup       text default ''    -- jour de récupération souhaité (ISO)
);

-- Index pour tri/filtre rapides
create index if not exists commandes_date_idx   on public.commandes (date);
create index if not exists commandes_etat_idx   on public.commandes (etat);
create index if not exists commandes_marque_idx on public.commandes (marque);

-- Met à jour updated_at automatiquement
create or replace function public.touch_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end $$;

drop trigger if exists commandes_touch on public.commandes;
create trigger commandes_touch
  before update on public.commandes
  for each row execute function public.touch_updated_at();

-- ────────────────────────────────────────────────────────────────────────────
--  Sécurité (RLS)
--  Outil interne magasin : on autorise lecture + écriture avec la clé "anon".
--  ⚠️  N'importe qui ayant l'URL + la clé anon peut lire/écrire.
--      Pour un usage strictement privé, ajoutez l'authentification Supabase
--      plus tard et remplacez les policies ci-dessous.
-- ────────────────────────────────────────────────────────────────────────────
alter table public.commandes enable row level security;

drop policy if exists "acces_complet_anon" on public.commandes;
create policy "acces_complet_anon"
  on public.commandes
  for all
  to anon, authenticated
  using (true)
  with check (true);
