import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import * as db from "./supabaseClient";
import { MAGASINS, MAGASIN_CODES, magasinNom, magasinTel } from "./magasins";

// ════════════════════════════════════════════════════════════════════════════
//   CONSTANTES — listes officielles (feuille 2 du tableau source)
// ════════════════════════════════════════════════════════════════════════════

// Fallback (utilisé en cas d'échec de chargement Supabase ou au tout début).
// Les vraies couleurs/noms sont chargés depuis Supabase et modifiables par les managers.
const ETATS_FALLBACK = [
  { val: "A commander",           label: "À commander",       color: "#E63946" },
  { val: "Rupture Stock",         label: "Rupture stock",     color: "#9B59B6" },
  { val: "DEPOT CLOANE",          label: "Dépôt CLOANE",      color: "#EC7BAA" },
  { val: "En Panier B2B",         label: "En panier B2B",     color: "#F39C12" },
  { val: "Commande validée",      label: "Commande validée",  color: "#8FCB87" },
  { val: "Commande réceptionnée", label: "Cde réceptionnée",  color: "#F1C40F" },
  { val: "Client prévenu",        label: "Client prévenu",    color: "#3498DB" },
  { val: "Produit encaissé",      label: "Produit encaissé",  color: "#2D7A4D" },
];

// Génère une couleur de fond pastel à partir d'une couleur hex principale
function lightenColor(hex, mix = 0.86) {
  if (!hex || typeof hex !== "string") return "#F5F0E8";
  const c = hex.replace("#", "");
  if (c.length !== 6) return "#F5F0E8";
  const r = parseInt(c.slice(0,2),16);
  const g = parseInt(c.slice(2,4),16);
  const b = parseInt(c.slice(4,6),16);
  const m = (v) => Math.round(v + (255 - v) * mix);
  const hh = (n) => n.toString(16).padStart(2,"0");
  return "#" + hh(m(r)) + hh(m(g)) + hh(m(b));
}

// Statuts "commande passée" → exemptés de la règle de purge à 60 jours
const STATUTS_COMMANDE = ["Commande validée", "Commande réceptionnée"];

const DEFAULT_VENDORS = ["ANTOINE","AURELIA","BLANDINE","CHARLOTTE","DJIBY","DORIAN","ESTELLE","MARGOT","MATTHIS","PACÔME"];

const DEFAULT_BRANDS = ["AIGLE","AMERICAN V.","ARTLOVE","BOMBERS","BSB","CABAIA","CALVIN KLEIN JEANS","CARHARTT","CHIC AU SOLEIL","CL11 SNEAKERS","COOLWAY","DAYTONA 73","DICKIES","EDWIN","FAGUO","FAM","FRED PERRY","FREEMAN T PORTER","GERTRUDE","GERRY ST TROPEZ","GIPSY","GRACE ET MILA","HERSCHEL","IZAC","JONSEN ISLAND","JOTT","KOST","KRAKATAU","LACOSTE","LEVIS","LA PETITE ETOILE","LA PETITE LAINE","LE TEMPS DES CERISES","LXH","MARIE ILE DE RÉ","NO NAME","THE NORTH FACE","NORTH SAILS","NORTH SAILS SHOES","OAKWOOD","OPTIMIST SUD FINITERE","ORBITKEY","PAKO LITTO","PULL IN","PYRENEX","QUARTIER IODE","REDSKINS","REPLAY","SALSA","SCHOTT","SCOTCH AND SODA","SCHMOOVE","SECRID","SIGNE NATURE","SWEET PANTS","TANTÄ","TOMMY HILFIGER","TOMMY HILFIGER JEANS","THE SURFCAR","VANESSA WU"];

const COLORS_TOP15 = ["Noir","Blanc","Gris","Beige","Marine","Bleu","Camel","Kaki","Marron","Vert","Rouge","Rose","Violet","Denim","Écru"];

const DAY_NAMES_FR = ["dimanche","lundi","mardi","mercredi","jeudi","vendredi","samedi"];

// ════════════════════════════════════════════════════════════════════════════
//   UTILITAIRES
// ════════════════════════════════════════════════════════════════════════════

function uid() { return Math.random().toString(36).slice(2,10); }

function parseDate(s) {
  if (!s) return null;
  if (s instanceof Date) return isNaN(s) ? null : s;
  s = String(s).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  return null;
}

function fmtDate(d) {
  if (!d) return "";
  if (typeof d === "string") d = parseDate(d);
  if (!d || isNaN(d)) return "";
  const yy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${yy}-${mm}-${dd}`;
}

function fmtDateFr(s) {
  const d = parseDate(s);
  if (!d) return "";
  return d.toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit",year:"numeric"});
}

function fmtDateShortFr(s) {
  const d = parseDate(s);
  if (!d) return "";
  return d.toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit"});
}

function addDays(date, n) {
  const d = date instanceof Date ? new Date(date) : parseDate(date);
  if (!d) return null;
  d.setDate(d.getDate() + n);
  return d;
}

function shiftMonths(dateStr, n) {
  const d = parseDate(dateStr);
  if (!d) return dateStr;
  d.setMonth(d.getMonth() + n);
  return fmtDate(d);
}

function dayName(d = new Date()) { return DAY_NAMES_FR[d.getDay()]; }

// Remet le 0 manquant : les n° clients commencent par 06 / 07 / 02.
// Beaucoup de numéros du tableau ont perdu leur 0 initial (9 chiffres).
function normalizePhone(t) {
  if (!t) return "";
  let s = String(t).replace(/[^\d]/g, "");
  if (!s) return "";
  // 9 chiffres commençant par 6/7/2 → il manque le 0
  if (s.length === 9 && /^[6729]/.test(s)) s = "0" + s;
  // déjà 33xxxxxxxxx → on repasse en 0xxxxxxxxx
  if (s.length === 11 && s.startsWith("33")) s = "0" + s.slice(2);
  return s;
}

function fmtPhone(t) {
  const s = normalizePhone(t);
  if (s.length === 10) return s.replace(/(\d{2})(?=\d)/g, "$1 ").trim();
  return s || String(t || "");
}

// Format international WhatsApp : 0612345678 → 33612345678
function phoneIntl(t) {
  const s = normalizePhone(t);
  if (s.length === 10 && s.startsWith("0")) return "33" + s.slice(1);
  return s;
}

// Message client prédéfini "commande arrivée"
// Logo CLOANE (noir sur transparent) — embed pour imprimer ticket caisse
const CLOANE_LOGO = "data:image/png;base64,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";

// Mise en forme automatique des saisies clavier
function upperCaseFr(s) { return (s || '').toString().toUpperCase(); }
function nameCase(s) {
  if (!s) return '';
  // Première lettre de chaque mot/partie (gère espaces, tirets, apostrophes)
  return s.toLowerCase().replace(/(^|[\s\-'\u2019])(\p{L})/gu, (_, sep, c) => sep + c.toUpperCase());
}

// Message "commande arrivée" — magasin = code (SQUARE/SIGNATURE/STORE)
function buildClientMessage(row, magasinCode) {
  const prenom = (row.prenom || "").trim() || "";
  const marque = (row.marque || "").trim();
  const m = MAGASINS[magasinCode] || MAGASINS.SQUARE;
  return (
`Bonjour ${prenom},

Votre commande${marque ? " de la marque " + marque : ""} est bien arrivée en magasin.

Merci de nous recontacter afin de confirmer votre futur passage pour la récupérer.

En vous souhaitant une bonne journée,
L'équipe ${m.nom}
${m.tel}`
  );
}

// Message "rupture de stock" — désolé, recontacter plus tard dans la saison
function buildRuptureMessage(row, magasinCode) {
  const prenom = (row.prenom || "").trim() || "";
  const marque = (row.marque || "").trim();
  const m = MAGASINS[magasinCode] || MAGASINS.SQUARE;
  return (
`Bonjour ${prenom},

Nous sommes désolés de vous informer que votre commande${marque ? " de la marque " + marque : ""} est malheureusement en rupture de stock chez notre fournisseur.

N'hésitez pas à nous recontacter plus tard dans la saison, au cas où ce produit reviendrait disponible.

Avec toutes nos excuses,
L'équipe ${m.nom}
${m.tel}`
  );
}

// ════════════════════════════════════════════════════════════════════════════
//   DONNÉES SOURCE — 471 lignes, dates décalées de +12 mois pour la démo
// ════════════════════════════════════════════════════════════════════════════

// (Données chargées depuis Supabase — voir supabaseClient.js)


const EMPTY_FORM = {
  date:"", nom:"", prenom:"", tel:"", email:"", marque:"", modele:"",
  refInt:"", refFourn:"", couleur:"", taille:"",
  etat:"", dateValid:"", vendeur:"", commentaire:"",
  articleCote:"", deCoteJusquau:""
};

// ── Helper components ──────────────────────────────────────────────────────

function badgeEtat(etat, statuts) {
  const s = (statuts || ETATS_FALLBACK).find(x => (x.val || x.nom) === etat);
  if (!s) return etat ? <span style={{fontSize:10,color:"#A09080"}}>{etat}</span> : null;
  const color = s.color || s.couleur;
  const bg = s.bg || lightenColor(color);
  const label = s.label || s.nom;
  return (
    <span style={{
      display:"inline-block",padding:"3px 9px",borderRadius:20,
      fontSize:10,fontWeight:600,fontFamily:"'DM Sans',sans-serif",
      color:color,background:bg,whiteSpace:"nowrap",letterSpacing:"0.02em",
    }}>{label}</span>
  );
}

function DropDown({options, value, onChange, placeholder, minWidth=140, allowEmpty=true}) {
  return (
    <select value={value||""} onChange={e=>onChange(e.target.value)}
      style={{
        fontSize:12,fontFamily:"'DM Sans',sans-serif",
        border:"1px solid #E0D8CE",borderRadius:7,
        padding:"5px 28px 5px 10px",background:"#FAFAF8",
        color:value?"#1C1510":"#A09080",cursor:"pointer",outline:"none",minWidth,
        appearance:"none",
        backgroundImage:`url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%238A7A6A'/%3E%3C/svg%3E")`,
        backgroundRepeat:"no-repeat",backgroundPosition:"right 8px center",
      }}>
      {allowEmpty && <option value="">{placeholder}</option>}
      {options.map(o => {
        const v = typeof o === "string" ? o : o.val;
        const l = typeof o === "string" ? o : (o.label || o.val);
        return <option key={v} value={v}>{l}</option>;
      })}
    </select>
  );
}

function FilterTag({label, onRemove}) {
  return (
    <span style={{
      display:"inline-flex",alignItems:"center",gap:6,
      background:"#F0EBE3",color:"#5A4030",
      fontSize:11,fontFamily:"'DM Sans',sans-serif",
      padding:"3px 8px 3px 12px",borderRadius:20,
      border:"1px solid #D8CCBE",
    }}>
      {label}
      <button onClick={onRemove} style={{
        background:"none",border:"none",cursor:"pointer",
        color:"#A09080",fontSize:14,lineHeight:1,padding:"0 0 0 2px",
      }}>×</button>
    </span>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//   MAIN COMPONENT
// ══════════════════════════════════════════════════════════════════════════

export default function CommandesB2B({ session, profile, onSignOut, onBack }) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(false);
  const [brands, setBrands] = useState([...DEFAULT_BRANDS]);
  const [vendeursAll, setVendeursAll] = useState([]); // [{id,nom,magasins:[]}]
  const [statutsAll, setStatutsAll] = useState([]); // [{id,nom,couleur,ordre}]

  // ETATS normalisés (compat avec l'ancien format {val,label,color,bg})
  const ETATS = useMemo(() => {
    const src = statutsAll.length ? statutsAll : ETATS_FALLBACK;
    return src.map(s => ({
      val: s.val || s.nom,
      label: s.label || s.nom,
      color: s.color || s.couleur || "#8A7A6A",
      bg: s.bg || lightenColor(s.color || s.couleur || "#8A7A6A"),
      id: s.id, // utile pour les paramètres
      ordre: s.ordre,
    }));
  }, [statutsAll]);
  // Infos B2B par fournisseur : URL site B2B, téléphone et e-mail du contact
  const [brandUrls, setBrandUrls] = useState({
    "CARHARTT":       { url: "https://b2b.carhartt-wip.com", tel: "", email: "" },
    "TOMMY HILFIGER": { url: "https://b2b.tommy.com",        tel: "", email: "" },
    "LACOSTE":        { url: "https://b2b.lacoste.com",      tel: "", email: "" },
    "AIGLE":          { url: "https://pro.aigle.com",        tel: "", email: "" },
    "REPLAY":         { url: "https://b2b.replay.it",        tel: "", email: "" },
  });
  const [filters, setFilters] = useState({etat:"",marque:"",vendeur:""});
  const [search, setSearch] = useState("");
  const [sortCol, setSortCol] = useState("date");
  const [sortDir, setSortDir] = useState("desc");
  const [bulkMode, setBulkMode] = useState(false);
  const [selected, setSelected] = useState(new Set());
  const [showEdit, setShowEdit] = useState(false);
  const [editId, setEditId] = useState(null);
  const [form, setForm] = useState(EMPTY_FORM);
  const [ticketRow, setTicketRow] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [settingsTab, setSettingsTab] = useState("brands");
  const [lastBackupDay, setLastBackupDay] = useState(null);
  const [expandedId, setExpandedId] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [archiveCount, setArchiveCount] = useState(0);
  const [newBrandInput, setNewBrandInput] = useState("");
  const [newVendorInput, setNewVendorInput] = useState("");
  const [contactMenu, setContactMenu] = useState(null); // {row, x, y} popover contact client
  const [statusMenu, setStatusMenu] = useState(null);   // {row, x, y} popover changement de statut
  const [ruptureRow, setRuptureRow] = useState(null);   // modal "rupture stock — prévenir client"

  // Multi-magasins : magasin actif (filtre)
  // null = tous les magasins (managers + vendeurs multi-magasins)
  const userMagasins = profile?.magasins || [];
  const isManager = profile?.role === "manager";
  const [activeMagasin, setActiveMagasin] = useState(
    userMagasins.length === 1 ? userMagasins[0] : null
  );

  // Vue archives (par défaut désactivé)
  const [showArchive, setShowArchive] = useState(false);

  // ── CHARGEMENT depuis Supabase + auto-archive au montage ───────────────────
  // Règles :
  //   • "Produit encaissé" plus vieux que 6 jours → passe en archive
  //   • Commande archivée depuis plus de 5 jours → suppression définitive
  //   • Commandes hors statut "commandé" plus vieilles que 60 jours → archive
  useEffect(() => {
    (async () => {
      try {
        const rows = await db.fetchOrders();
        const today = new Date();
        const sixDaysAgo = addDays(today, -6);
        const sixtyDaysAgo = addDays(today, -60);
        const fiveDaysAgo = new Date(today.getTime() - 5*86400000);

        // 1. Suppression définitive : archivées depuis +5 jours
        const toDelete = rows
          .filter(r => r.archivedAt && new Date(r.archivedAt) < fiveDaysAgo)
          .map(r => r.id);
        if (toDelete.length) db.deleteOrders(toDelete).catch(e => console.error("delete archived", e));

        // 2. Auto-archive : "Produit encaissé" de +6 jours, et anciennes commandes de +60 jours
        const toArchive = rows
          .filter(r => !r.archivedAt) // pas déjà archivée
          .filter(r => {
            if (r.etat === "Produit encaissé") {
              const dv = parseDate(r.dateValid) || parseDate(r.date);
              return dv && dv < sixDaysAgo;
            }
            if (!STATUTS_COMMANDE.includes(r.etat)) {
              const dc = parseDate(r.date);
              return dc && dc < sixtyDaysAgo;
            }
            return false;
          })
          .map(r => r.id);

        let archivedNow = null;
        if (toArchive.length) {
          try { archivedNow = await db.archiveOrders(toArchive); }
          catch (e) { console.error("archive", e); }
        }

        // 3. Mettre à jour la liste locale
        const cleaned = rows
          .filter(r => !toDelete.includes(r.id))
          .map(r => toArchive.includes(r.id) ? {...r, archivedAt: archivedNow || new Date().toISOString()} : r);

        setData(cleaned);
        setArchiveCount(toArchive.length + toDelete.length);
      } catch (e) {
        console.error("Chargement Supabase échoué :", e);
        setLoadError(true);
      } finally {
        setLoading(false);
        setTimeout(() => setLoaded(true), 60);
      }
      // Charger la liste des vendeurs depuis Supabase
      try {
        const v = await db.fetchVendeurs();
        setVendeursAll(v);
      } catch (e) { console.error("fetchVendeurs", e); }
      // Charger la liste des statuts depuis Supabase
      try {
        const s = await db.fetchStatuts();
        setStatutsAll(s);
      } catch (e) { console.error("fetchStatuts", e); }
    })();
  }, []);

  // Liste des noms de vendeurs visibles selon le magasin actif
  // (si "Tous" : on affiche tous les vendeurs des magasins autorisés)
  const visibleVendors = useMemo(() => {
    const filterMagasins = activeMagasin ? [activeMagasin] : userMagasins;
    return vendeursAll
      .filter(v => v.magasins.some(m => filterMagasins.includes(m)))
      .map(v => v.nom);
  }, [vendeursAll, activeMagasin, userMagasins]);

  // ── Filtering & sorting ──
  const filtered = useMemo(() => {
    let d = data;
    // 1. Filtre par magasin actif (si "Tous" → garde les magasins autorisés)
    if (activeMagasin) {
      d = d.filter(r => r.magasin === activeMagasin);
    } else {
      d = d.filter(r => userMagasins.includes(r.magasin));
    }
    // 2. Onglet : Actives ou Archivées
    d = d.filter(r => showArchive ? !!r.archivedAt : !r.archivedAt);
    // 3. Filtres existants
    if (filters.etat) d = d.filter(r => r.etat === filters.etat);
    if (filters.marque) d = d.filter(r => r.marque === filters.marque);
    if (filters.vendeur) d = d.filter(r => r.vendeur === filters.vendeur);
    if (search.trim()) {
      const s = search.toLowerCase();
      d = d.filter(r =>
        r.nom.toLowerCase().includes(s) ||
        r.prenom.toLowerCase().includes(s) ||
        r.marque.toLowerCase().includes(s) ||
        r.modele.toLowerCase().includes(s) ||
        String(r.tel).includes(s) ||
        r.couleur.toLowerCase().includes(s) ||
        r.refInt.toLowerCase().includes(s) ||
        r.refFourn.toLowerCase().includes(s)
      );
    }
    return [...d].sort((a,b) => {
      const va = a[sortCol]||""; const vb = b[sortCol]||"";
      if (sortCol === "date" || sortCol === "dateValid") {
        const da = parseDate(va); const db = parseDate(vb);
        if (!da && !db) return 0;
        if (!da) return 1; if (!db) return -1;
        return sortDir === "asc" ? da-db : db-da;
      }
      return sortDir === "asc"
        ? String(va).localeCompare(String(vb))
        : String(vb).localeCompare(String(va));
    });
  }, [data, filters, search, sortCol, sortDir]);

  // Données filtrées par magasin et vue archive (avant filtres etat/marque/vendeur)
  // utilisées pour calculer les stats et le compteur des onglets de manière cohérente
  const magasinFiltered = useMemo(() => {
    let d = data;
    if (activeMagasin) d = d.filter(r => r.magasin === activeMagasin);
    else d = d.filter(r => userMagasins.includes(r.magasin));
    return d;
  }, [data, activeMagasin, userMagasins]);

  const activesData = useMemo(() => magasinFiltered.filter(r => !r.archivedAt), [magasinFiltered]);
  const archivedData = useMemo(() => magasinFiltered.filter(r => !!r.archivedAt), [magasinFiltered]);

  // Stats reflètent la vue active (magasin + onglet)
  const statsBase = showArchive ? archivedData : activesData;
  const stats = useMemo(() => ({
    total: statsBase.length,
    aCommander: statsBase.filter(r => r.etat === "A commander").length,
    clientPrevenu: statsBase.filter(r => r.etat === "Client prévenu").length,
    rupture: statsBase.filter(r => r.etat === "Rupture Stock").length,
    encaisse: statsBase.filter(r => r.etat === "Produit encaissé").length,
  }), [statsBase]);

  const todayDayName = useMemo(() => dayName(), []);
  const backupAvailable = todayDayName !== "dimanche" && lastBackupDay !== todayDayName;

  // ── Handlers ──
  function toggleSort(col) {
    if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  }

  function updateRow(id, patch) {
    setData(d => d.map(r => r.id === id ? {...r, ...patch} : r));
    db.updateOrder(id, patch).catch(e => console.error("updateOrder", e));
  }

  function changeStatus(rowId, newStatus) {
    const today = fmtDate(new Date());
    const row = data.find(r => r.id === rowId);
    if (!row) return;
    updateRow(rowId, {etat: newStatus, dateValid: today});
    if (newStatus === "Client prévenu") {
      setTicketRow({
        ...row, etat: newStatus, dateValid: today,
        articleCote: row.articleCote || [row.marque, row.modele, row.couleur && "Couleur : "+row.couleur, row.taille && "Taille : "+row.taille].filter(Boolean).join("\n"),
        deCoteJusquau: row.deCoteJusquau || fmtDate(addDays(new Date(), 7)),
      });
    }
  }

  // Suppression manuelle = archivage (visible 5 jours dans l'onglet Archivé puis suppression auto)
  // Sauf si déjà archivé : auquel cas suppression définitive
  function deleteRow(id) {
    const row = data.find(r => r.id === id);
    if (!row) return;
    const isArchived = !!row.archivedAt;
    const msg = isArchived
      ? "Supprimer DÉFINITIVEMENT cette commande archivée ?"
      : "Archiver cette commande ? Elle restera visible 5 jours dans l'onglet Archivé puis sera supprimée automatiquement.";
    if (!confirm(msg)) return;
    if (isArchived) {
      setData(d => d.filter(r => r.id !== id));
      setSelected(s => { const n = new Set(s); n.delete(id); return n; });
      db.deleteOrder(id).catch(e => console.error("deleteOrder", e));
    } else {
      const now = new Date().toISOString();
      setData(d => d.map(r => r.id === id ? {...r, archivedAt: now} : r));
      setSelected(s => { const n = new Set(s); n.delete(id); return n; });
      db.archiveOrders([id]).catch(e => console.error("archive", e));
    }
  }

  function openNew() {
    setForm({...EMPTY_FORM, date: fmtDate(new Date())});
    setEditId(null);
    setShowEdit(true);
  }

  function openEdit(row) {
    setForm({...row});
    setEditId(row.id);
    setShowEdit(true);
  }

  async function saveForm() {
    if (!form.nom?.trim()) {
      alert("Le nom du client est obligatoire.");
      return;
    }
    if (!form.tel?.trim()) {
      alert("Le numéro de téléphone est obligatoire.");
      return;
    }
    const today = fmtDate(new Date());
    if (editId) {
      const oldRow = data.find(r => r.id === editId);
      const statusChanged = oldRow && oldRow.etat !== form.etat;
      const newDateValid = statusChanged ? today : form.dateValid;
      const merged = {...form, dateValid: newDateValid, magasin: form.magasin || oldRow?.magasin};
      updateRow(editId, merged);
      if (statusChanged && form.etat === "Client prévenu") {
        setTimeout(() => setTicketRow({
          ...merged, id: editId,
          articleCote: form.articleCote || [form.marque, form.modele, form.couleur && "Couleur : "+form.couleur, form.taille && "Taille : "+form.taille].filter(Boolean).join("\n"),
          deCoteJusquau: form.deCoteJusquau || fmtDate(addDays(new Date(), 7)),
        }), 200);
      }
      if (statusChanged && form.etat === "Rupture Stock") {
        setTimeout(() => setRuptureRow({...merged, id: editId}), 200);
      }
      setShowEdit(false);
    } else {
      // Magasin par défaut : le magasin actif, sinon le premier magasin du profil
      const magasin = form.magasin || activeMagasin || userMagasins[0] || "SQUARE";
      const draft = {...form, magasin, date: form.date || today};
      setShowEdit(false);
      try {
        const newRow = await db.insertOrder(draft);
        setData(d => [newRow, ...d]);
        if (newRow.etat === "Client prévenu") {
          setTimeout(() => setTicketRow({
            ...newRow,
            articleCote: newRow.articleCote || [newRow.marque, newRow.modele, newRow.couleur && "Couleur : "+newRow.couleur, newRow.taille && "Taille : "+newRow.taille].filter(Boolean).join("\n"),
            deCoteJusquau: newRow.deCoteJusquau || fmtDate(addDays(new Date(), 7)),
          }), 200);
        }
        if (newRow.etat === "Rupture Stock") {
          setTimeout(() => setRuptureRow(newRow), 200);
        }
      } catch (e) {
        console.error("insertOrder", e);
        alert("Erreur lors de l'enregistrement en ligne. Vérifiez la connexion Supabase.");
      }
    }
  }

  // ── Bulk selection handlers ──
  // Le bouton fait apparaître/disparaître les cases à cocher SANS tout cocher.
  function toggleSelectMode() {
    if (!bulkMode) {
      setBulkMode(true);
      setSelected(new Set());
    } else {
      setBulkMode(false);
      setSelected(new Set());
    }
  }
  // Coche / décoche toutes les lignes visibles (case d'en-tête)
  function toggleSelectAllRows() {
    if (selected.size === filtered.length) setSelected(new Set());
    else setSelected(new Set(filtered.map(r => r.id)));
  }

  function toggleSelect(id) {
    setSelected(s => {
      const n = new Set(s);
      n.has(id) ? n.delete(id) : n.add(id);
      return n;
    });
  }

  function exitBulkMode() {
    setBulkMode(false);
    setSelected(new Set());
  }

  function applyBulkStatus(newStatus) {
    if (!newStatus || selected.size === 0) return;
    const today = fmtDate(new Date());
    const ids = [...selected];
    setData(d => d.map(r => selected.has(r.id) ? {...r, etat: newStatus, dateValid: today} : r));
    const count = ids.length;
    setSelected(new Set());
    setBulkMode(false);
    db.bulkUpdateStatus(ids, newStatus, today).catch(e => console.error("bulkUpdate", e));
    if (newStatus === "Client prévenu") {
      setTimeout(() => alert(`${count} commande(s) passée(s) en « Client prévenu ». Cliquez sur l'icône 🎫 dans chaque ligne pour générer le ticket caisse correspondant.`), 100);
    }
  }

  // ── Ticket save ──
  function saveTicket() {
    if (!ticketRow) return;
    updateRow(ticketRow.id, {
      articleCote: ticketRow.articleCote,
      deCoteJusquau: ticketRow.deCoteJusquau,
      vendeur: ticketRow.vendeur,
    });
  }

  // ── Brand / Vendor management ──
  function addBrand() {
    const name = newBrandInput.trim().toUpperCase();
    if (!name || brands.includes(name)) return;
    setBrands(b => [...b, name].sort());
    setNewBrandInput("");
  }
  function removeBrand(name) {
    if (!confirm(`Supprimer la marque "${name}" de la liste ?`)) return;
    setBrands(b => b.filter(x => x !== name));
    setBrandUrls(u => { const n = {...u}; delete n[name]; return n; });
  }
  // Ouvre le site B2B de la marque (bouton "Commander")
  // Bouton "Commander" : ouvre le site B2B si URL renseignée,
  // sinon ouvre un mail prérempli au fournisseur si email renseigné,
  // sinon propose d'aller renseigner l'un ou l'autre.
  function openOrderUrl(row) {
    // Supporte les 2 signatures : openOrderUrl(row) ou openOrderUrl(marqueString)
    const marque = typeof row === "string" ? row : row?.marque;
    if (!marque) return;
    const info = brandUrls[marque] || {};
    const url = (info.url || "").trim();
    const email = (info.email || "").trim();

    if (url) {
      const full = /^https?:\/\//i.test(url) ? url : "https://" + url;
      window.open(full, "_blank", "noopener");
      return;
    }
    if (email) {
      const m = MAGASINS[row?.magasin || activeMagasin || "SQUARE"] || MAGASINS.SQUARE;
      const subject = encodeURIComponent(`Commande client — ${marque} — ${m.nom}`);
      const body = encodeURIComponent(buildSupplierOrderMessage(row, marque, m));
      window.location.href = `mailto:${email}?subject=${subject}&body=${body}`;
      return;
    }
    alert(
      `Aucun lien B2B ni e-mail enregistrés pour « ${marque} ».\n` +
      `Renseignez au moins l'un des deux dans ⚙ Paramètres → Fournisseurs.`
    );
    setShowSettings(true); setSettingsTab("brands");
  }
  // Construit le mail au fournisseur avec le détail de la commande
  function buildSupplierOrderMessage(row, marque, magasin) {
    const lines = ["Bonjour,", "", "Pourriez-vous nous expédier l'article suivant pour une commande client :", ""];
    lines.push(`Marque : ${marque}`);
    if (row?.modele)    lines.push(`Modèle : ${row.modele}`);
    if (row?.refInt)    lines.push(`Référence interne : ${row.refInt}`);
    if (row?.refFourn)  lines.push(`Référence fournisseur : ${row.refFourn}`);
    if (row?.couleur && row.couleur.trim()) lines.push(`Couleur : ${row.couleur.trim()}`);
    if (row?.taille)    lines.push(`Taille : ${row.taille}`);
    lines.push("", "Merci d'avance,", `L'équipe ${magasin.nom}`, magasin.tel);
    return lines.join("\n");
  }
  // Met à jour une info fournisseur (url, tel ou email)
  function setBrandField(name, field, value) {
    setBrandUrls(u => ({
      ...u,
      [name]: { url:"", tel:"", email:"", ...(u[name]||{}), [field]: value }
    }));
  }
  // Contact client : WhatsApp / Mail / SMS avec message + tracker boolean
  function sendWhatsApp(row) {
    const intl = phoneIntl(row.tel);
    if (!intl) { alert("Numéro de téléphone manquant."); return; }
    const msg = encodeURIComponent(buildClientMessage(row, row.magasin));
    window.open(`https://wa.me/${intl}?text=${msg}`, "_blank", "noopener");
    updateRow(row.id, { msgWhatsapp: true });
    setContactMenu(null);
  }
  function sendMail(row) {
    if (!row.email) { alert("Aucune adresse e-mail renseignée pour ce client (fiche commande)."); return; }
    const subject = encodeURIComponent("Votre commande CLOANE est arrivée");
    const body = encodeURIComponent(buildClientMessage(row, row.magasin));
    window.location.href = `mailto:${row.email}?subject=${subject}&body=${body}`;
    updateRow(row.id, { mailEnvoye: true });
    setContactMenu(null);
  }
  function sendSMS(row) {
    const tel = normalizePhone(row.tel);
    if (!tel) { alert("Numéro de téléphone manquant."); return; }
    const body = encodeURIComponent(buildClientMessage(row, row.magasin));
    window.location.href = `sms:${tel}?&body=${body}`;
    setContactMenu(null);
  }
  // Mêmes fonctions mais pour le message "rupture de stock"
  function sendRuptureWhatsApp(row) {
    const intl = phoneIntl(row.tel);
    if (!intl) { alert("Numéro de téléphone manquant."); return; }
    const msg = encodeURIComponent(buildRuptureMessage(row, row.magasin));
    window.open(`https://wa.me/${intl}?text=${msg}`, "_blank", "noopener");
    updateRow(row.id, { msgWhatsapp: true });
  }
  function sendRuptureMail(row) {
    if (!row.email) { alert("Aucune adresse e-mail renseignée pour ce client."); return; }
    const subject = encodeURIComponent("Votre commande CLOANE");
    const body = encodeURIComponent(buildRuptureMessage(row, row.magasin));
    window.location.href = `mailto:${row.email}?subject=${subject}&body=${body}`;
    updateRow(row.id, { mailEnvoye: true });
  }
  function sendRuptureSMS(row) {
    const tel = normalizePhone(row.tel);
    if (!tel) { alert("Numéro de téléphone manquant."); return; }
    const body = encodeURIComponent(buildRuptureMessage(row, row.magasin));
    window.location.href = `sms:${tel}?&body=${body}`;
  }
  async function addVendor() {
    const name = upperCaseFr(newVendorInput.trim());
    if (!name) return;
    if (vendeursAll.some(v => v.nom === name)) {
      alert(`Le conseiller "${name}" existe déjà.`);
      return;
    }
    // Par défaut : affecté au magasin actif (ou au premier magasin du profil)
    const magasins = activeMagasin ? [activeMagasin] : (userMagasins.length === 1 ? [userMagasins[0]] : []);
    try {
      const created = await db.insertVendeur(name, magasins);
      setVendeursAll(v => [...v, created].sort((a,b) => a.nom.localeCompare(b.nom)));
      setNewVendorInput("");
    } catch (e) {
      console.error(e);
      alert("Impossible d'ajouter ce conseiller. Vérifiez vos droits.");
    }
  }
  async function removeVendeur(id, name) {
    if (!confirm(`Supprimer le conseiller "${name}" de la liste ?`)) return;
    setVendeursAll(v => v.filter(x => x.id !== id));
    try { await db.deleteVendeur(id); }
    catch (e) {
      console.error(e);
      alert("Impossible de supprimer ce conseiller. Vérifiez vos droits.");
      // Recharger pour resync
      db.fetchVendeurs().then(setVendeursAll).catch(()=>{});
    }
  }
  async function toggleVendeurMagasin(vendeur, magasinCode) {
    const newMagasins = vendeur.magasins.includes(magasinCode)
      ? vendeur.magasins.filter(m => m !== magasinCode)
      : [...vendeur.magasins, magasinCode];
    setVendeursAll(arr => arr.map(v => v.id === vendeur.id ? {...v, magasins: newMagasins} : v));
    try { await db.updateVendeurMagasins(vendeur.id, newMagasins); }
    catch (e) {
      console.error(e);
      alert("Impossible de modifier l'affectation. Vérifiez vos droits.");
      db.fetchVendeurs().then(setVendeursAll).catch(()=>{});
    }
  }

  // ── Statuts CRUD ──
  async function updateStatutColor(id, couleur) {
    setStatutsAll(arr => arr.map(s => s.id === id ? {...s, couleur} : s));
    try { await db.updateStatut(id, { couleur }); }
    catch (e) {
      console.error(e);
      alert("Impossible de modifier la couleur. Vérifiez vos droits.");
      db.fetchStatuts().then(setStatutsAll).catch(()=>{});
    }
  }
  // Saisie en cours (édition contrôlée locale)
  function renameStatutInline(id, oldNom, newNom) {
    setStatutsAll(arr => arr.map(s => s.id === id ? {...s, nom: newNom} : s));
  }
  // Validation du renommage (au blur) : cascade en base
  async function commitStatutRename(id, oldNom, newNom) {
    const trimmed = (newNom || "").trim();
    if (!trimmed || trimmed === oldNom) {
      // Resync si vide
      if (!trimmed) {
        setStatutsAll(arr => arr.map(s => s.id === id ? {...s, nom: oldNom} : s));
      }
      return;
    }
    try {
      await db.renameStatut(oldNom, trimmed);
      // Reflet local : les commandes utilisant l'ancien nom passent au nouveau
      setData(arr => arr.map(r => r.etat === oldNom ? {...r, etat: trimmed} : r));
    } catch (e) {
      console.error(e);
      alert("Impossible de renommer ce statut. Vérifiez vos droits.");
      setStatutsAll(arr => arr.map(s => s.id === id ? {...s, nom: oldNom} : s));
    }
  }
  async function removeStatut(s) {
    const used = data.filter(r => r.etat === s.nom).length;
    if (used > 0) {
      alert(`Impossible de supprimer « ${s.nom} » : ce statut est utilisé par ${used} commande(s).`);
      return;
    }
    if (!confirm(`Supprimer définitivement le statut « ${s.nom} » ?`)) return;
    setStatutsAll(arr => arr.filter(x => x.id !== s.id));
    try { await db.deleteStatut(s.id); }
    catch (e) {
      console.error(e);
      alert("Impossible de supprimer ce statut. Vérifiez vos droits.");
      db.fetchStatuts().then(setStatutsAll).catch(()=>{});
    }
  }
  async function addStatut() {
    const nom = prompt("Nom du nouveau statut :");
    if (!nom?.trim()) return;
    const cleanNom = nom.trim();
    if (statutsAll.some(s => s.nom === cleanNom)) {
      alert("Ce statut existe déjà.");
      return;
    }
    const maxOrdre = Math.max(0, ...statutsAll.map(s => s.ordre || 0));
    try {
      const created = await db.insertStatut(cleanNom, "#8A7A6A", maxOrdre + 1);
      setStatutsAll(arr => [...arr, created].sort((a,b) => a.ordre - b.ordre));
    } catch (e) {
      console.error(e);
      alert("Impossible d'ajouter ce statut. Vérifiez vos droits.");
    }
  }

  // ── Daily Excel backup ──
  function exportBackup() {
    const day = dayName();
    if (day === "dimanche") {
      alert("Pas de sauvegarde le dimanche. Les sauvegardes s'effectuent du lundi au samedi.");
      return;
    }
    const rows = data.map(r => ({
      "Date création": fmtDateFr(r.date),
      "Nom": r.nom, "Prénom": r.prenom,
      "Téléphone": fmtPhone(r.tel), "E-mail": r.email || "",
      "Marque": r.marque, "Nom modèle": r.modele,
      "Ref. interne": r.refInt, "Ref. fournisseur": r.refFourn,
      "Couleur": (r.couleur||"").trim(), "Taille": r.taille,
      "Etat": r.etat, "Date validation": fmtDateFr(r.dateValid),
      "VENDEUR": r.vendeur,
      "Article de côté": r.articleCote,
      "De côté jusqu'au": fmtDateFr(r.deCoteJusquau),
      "Commentaire": r.commentaire,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Commandes");
    XLSX.writeFile(wb, `CLOANE_Commandes.${day}.xlsx`);
    setLastBackupDay(day);
  }

  function printTicket() {
    window.print();
  }

  const hasFilters = filters.etat || filters.marque || filters.vendeur;
  const allSelected = filtered.length > 0 && selected.size === filtered.length;

  // ── Inline style constants ──
  const TH_STYLE = {
    fontFamily:"'DM Sans',sans-serif",fontSize:10,fontWeight:700,
    color:"#8A7A6A",letterSpacing:"0.1em",textTransform:"uppercase",
    padding:"10px 12px",background:"#F5F0E8",
    borderBottom:"1px solid #E8E0D5",whiteSpace:"nowrap",
    cursor:"pointer",userSelect:"none",textAlign:"left",
  };
  const TD_STYLE = {
    padding:"9px 12px",borderBottom:"1px solid #F0EAE0",
    fontFamily:"'DM Sans',sans-serif",fontSize:12,color:"#2C1E0F",verticalAlign:"top",
  };

  return (
    <div style={{
      opacity: loaded?1:0,
      transform: loaded?"translateY(0)":"translateY(16px)",
      transition:"opacity .4s ease, transform .4s ease",
      background:"#F4F0E8",minHeight:"100vh",
      paddingBottom: bulkMode && selected.size > 0 ? 90 : 0,
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;400;500&family=DM+Sans:wght@300;400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        .tr-hover:hover { background:#FFF8F2 !important; }
        .row-selected { background:#FFF4E5 !important; }
        select:focus, input:focus, textarea:focus { outline:2px solid #C8A96E; outline-offset:0; }
        ::-webkit-scrollbar { height:6px; width:6px; }
        ::-webkit-scrollbar-track { background:#F0EBE3; }
        ::-webkit-scrollbar-thumb { background:#C8B89A; border-radius:10px; }
        @keyframes blink { 0%,100% { opacity:1 } 50% { opacity:.5 } }
        .pulse { animation: blink 1.6s ease-in-out infinite; }
        .ticket-paper {
          font-family:'Courier New','Courier',monospace;
          background:#FFFCF5; color:#1C1510;
          padding:20px 18px; line-height:1.55;
          border:1px dashed #D8CCBE; border-radius:6px;
          width:280px; margin:0 auto;
          font-weight:700;
        }
        @media (max-width: 640px) {
          .hide-on-mobile { display:none !important; }
        }
        @media print {
          body * { visibility:hidden !important; }
          .ticket-printable, .ticket-printable * { visibility:visible !important; }
          .ticket-printable {
            position:fixed !important; top:0 !important; left:0 !important;
            width:76mm !important; padding:4mm 3mm !important;
            border:none !important; background:#fff !important;
          }
          .ticket-paper {
            width:100% !important; border:none !important; padding:0 !important;
            margin:0 !important; background:#fff !important;
          }
          @page { size:76mm auto; margin:0; }
        }
      `}</style>

      {/* ── HEADER ─────────────────────────────────────────────────────── */}
      <div style={{
        position:"sticky",top:0,zIndex:50,
        background:"rgba(244,240,232,0.96)",backdropFilter:"blur(12px)",
        borderBottom:"1px solid rgba(200,169,110,0.2)",
        padding:"0 16px",display:"flex",alignItems:"center",
        justifyContent:"space-between",minHeight:60,gap:8,flexWrap:"wrap",
      }}>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",minWidth:0}}>
          <div style={{
            width:30,height:30,background:"#1C1510",borderRadius:"50%",
            display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,
          }}>
            <span style={{color:"#C8A96E",fontFamily:"'Cormorant Garamond',serif",fontSize:13,fontWeight:500}}>C</span>
          </div>
          <div style={{
            fontFamily:"'Cormorant Garamond',serif",fontSize:17,fontWeight:400,
            color:"#1C1510",letterSpacing:"0.06em",textTransform:"uppercase",
          }}>Commandes B2B</div>

          {/* Sélecteur magasin (si l'utilisateur a accès à plusieurs) */}
          {userMagasins.length > 1 && (
            <select value={activeMagasin || ""}
              onChange={e=>setActiveMagasin(e.target.value || null)}
              style={{
                fontSize:11,padding:"5px 10px",border:"1px solid #C8A96E",
                borderRadius:8,background:"#FFF8EB",color:"#1C1510",
                fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer",
                letterSpacing:"0.04em",
              }}>
              <option value="">🏪 Tous magasins</option>
              {userMagasins.map(m => (
                <option key={m} value={m}>{MAGASINS[m]?.nom || m}</option>
              ))}
            </select>
          )}
          {userMagasins.length === 1 && (
            <span style={{
              fontSize:11,padding:"4px 10px",background:"#FFF8EB",border:"1px solid #C8A96E",
              borderRadius:8,color:"#7A5A2A",fontFamily:"'DM Sans',sans-serif",fontWeight:600,
              letterSpacing:"0.04em",
            }}>{MAGASINS[userMagasins[0]]?.nom || userMagasins[0]}</span>
          )}

          <span style={{
            fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"#A09080",
            background:"#EDE4D5",padding:"2px 8px",borderRadius:20,
          }}>{filtered.length} / {data.length}</span>
        </div>

        <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
          {backupAvailable && (
            <button onClick={exportBackup} className="pulse" style={{
              background:"#FFF8EB",border:"1px solid #C8A96E",borderRadius:8,
              padding:"6px 11px",fontSize:11,cursor:"pointer",
              fontFamily:"'DM Sans',sans-serif",color:"#7A5A2A",
              display:"flex",alignItems:"center",gap:5,fontWeight:600,
            }} title={`Sauvegarde Excel quotidienne — ${todayDayName}`}>
              <span style={{fontSize:14}}>💾</span>
              <span className="hide-on-mobile">Sauvegarde</span>
            </button>
          )}
          <button onClick={()=>setShowSettings(true)} title="Paramètres" style={{
            background:"#F5F0E8",border:"1px solid #DDD4C8",borderRadius:10,
            padding:"7px 11px",fontSize:18,cursor:"pointer",color:"#5A4030",
            lineHeight:1,minHeight:36,display:"flex",alignItems:"center",justifyContent:"center",
          }}>⚙</button>
          <button onClick={openNew} title="Nouvelle commande" style={{
            background:"#1C1510",color:"#E8DED0",border:"none",borderRadius:10,
            padding:"8px 12px",fontSize:13,cursor:"pointer",
            fontFamily:"'DM Sans',sans-serif",letterSpacing:"0.04em",
            display:"flex",alignItems:"center",gap:6,fontWeight:600,minHeight:36,
          }}>
            <span style={{fontSize:18,lineHeight:0.8,fontWeight:300}}>+</span>
            <span className="hide-on-mobile">Nouvelle</span>
          </button>
          <button onClick={onSignOut} title={profile?.display_name || profile?.email} style={{
            background:"#FFF",border:"1px solid #DDD4C8",borderRadius:10,
            padding:"6px 10px",fontSize:14,cursor:"pointer",color:"#8A7A6A",
            lineHeight:1,minHeight:36,display:"flex",alignItems:"center",justifyContent:"center",
          }}>⎋</button>
        </div>
      </div>

      <div style={{padding:"18px 24px 32px"}}>

        {/* Loading / error */}
        {loading && (
          <div style={{
            background:"#FFF",border:"1px solid #EDE4D5",borderRadius:10,
            padding:"14px 18px",marginBottom:14,fontSize:12,
            fontFamily:"'DM Sans',sans-serif",color:"#8A7A6A",
          }}>⏳ Chargement des commandes depuis Supabase…</div>
        )}
        {loadError && (
          <div style={{
            background:"#FCEAEA",border:"1px solid #F0C8C8",borderRadius:10,
            padding:"14px 18px",marginBottom:14,fontSize:12,
            fontFamily:"'DM Sans',sans-serif",color:"#9B2020",
          }}>
            ⚠️ Impossible de joindre Supabase. Vérifiez vos variables d'environnement
            (<code>VITE_SUPABASE_URL</code> / <code>VITE_SUPABASE_ANON_KEY</code>) et que le schéma a bien été créé.
          </div>
        )}

        {/* Cleanup notice */}
        {archiveCount > 0 && (
          <div style={{
            background:"#F8F3EB",border:"1px solid #E8DCC0",borderRadius:10,
            padding:"8px 14px",marginBottom:14,fontSize:11,
            fontFamily:"'DM Sans',sans-serif",color:"#7A5A2A",
            display:"flex",alignItems:"center",gap:8,
          }}>
            <span>♻️</span>
            <span><b>{archiveCount}</b> commande(s) traitée(s) automatiquement (archivage des encaissés +6 jours, et suppression des archives +5 jours).</span>
          </div>
        )}

        {/* ── ONGLETS : Actives / Archivées ─────────────────────────────── */}
        <div style={{
          display:"flex",gap:4,marginBottom:18,
          background:"#F5F0E8",borderRadius:10,padding:4,
          width:"fit-content",
        }}>
          {[
            {key:false, label:"📋 Actives", count: activesData.length},
            {key:true,  label:"📦 Archivées", count: archivedData.length},
          ].map(t => (
            <button key={String(t.key)} onClick={()=>setShowArchive(t.key)} style={{
              background: showArchive===t.key ? "#1C1510" : "transparent",
              color: showArchive===t.key ? "#E8DED0" : "#5A4030",
              border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,
              fontFamily:"'DM Sans',sans-serif",fontWeight:600,cursor:"pointer",
              display:"flex",alignItems:"center",gap:6,letterSpacing:"0.04em",
            }}>
              {t.label}
              <span style={{
                background: showArchive===t.key ? "rgba(232,222,208,0.2)" : "#FFF",
                padding:"1px 8px",borderRadius:10,fontSize:10,fontWeight:700,
              }}>{t.count}</span>
            </button>
          ))}
        </div>

        {/* ── STATS ROW ──────────────────────────────────────────────── */}
        <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap"}}>
          {[
            {label:"Total",val:stats.total,etat:null},
            {label:"À commander",val:stats.aCommander,etat:"A commander"},
            {label:"Client prévenu",val:stats.clientPrevenu,etat:"Client prévenu"},
            {label:"Rupture stock",val:stats.rupture,etat:"Rupture Stock"},
            {label:"Encaissé",val:stats.encaisse,etat:"Produit encaissé"},
          ].map(s => {
            const meta = s.etat ? ETATS.find(e => e.val === s.etat) : null;
            const color = meta?.color || "#8A7A6A";
            const bg = meta?.bg || "#F5F0E8";
            return ({...s, color, bg});
          }).map(s => (
            <div key={s.label} onClick={()=>s.etat && setFilters(f => ({...f, etat: f.etat===s.etat?"":s.etat}))}
              style={{
                background:"#FFF",border:`1px solid ${s.bg}`,
                borderTop:`3px solid ${s.color}`,borderRadius:10,
                padding:"10px 16px",minWidth:104,
                cursor:s.etat?"pointer":"default",
                opacity: filters.etat && filters.etat!==s.etat ? 0.5 : 1,
                transition:"opacity .15s",
              }}>
              <div style={{fontSize:22,fontWeight:300,color:s.color,fontFamily:"'Cormorant Garamond',serif",lineHeight:1}}>{s.val}</div>
              <div style={{fontSize:10,color:"#8A7A6A",fontFamily:"'DM Sans',sans-serif",marginTop:3,letterSpacing:"0.04em"}}>{s.label}</div>
            </div>
          ))}
        </div>

        {/* ── FILTER BAR ─────────────────────────────────────────────── */}
        <div style={{
          background:"#FFF",border:"1px solid #EDE4D5",borderRadius:12,
          padding:"12px 16px",marginBottom:12,
          display:"flex",flexWrap:"wrap",gap:10,alignItems:"center",
        }}>
          <div style={{position:"relative",flex:"1 1 200px",minWidth:180}}>
            <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:13,color:"#A09080"}}>🔍</span>
            <input value={search} onChange={e=>setSearch(e.target.value)}
              placeholder="Rechercher client, marque, modèle, téléphone…"
              style={{
                width:"100%",padding:"7px 10px 7px 32px",
                fontSize:12,fontFamily:"'DM Sans',sans-serif",
                border:"1px solid #E0D8CE",borderRadius:7,
                background:"#FAFAF8",color:"#1C1510",
              }}/>
          </div>
          <div style={{width:1,height:24,background:"#E8E0D5"}}/>
          <span style={{fontSize:10,color:"#A09080",fontFamily:"'DM Sans',sans-serif",letterSpacing:"0.1em",textTransform:"uppercase",fontWeight:600}}>Filtres :</span>
          <DropDown options={ETATS} value={filters.etat}
            onChange={v=>setFilters(f=>({...f,etat:v}))}
            placeholder="Tous les statuts" minWidth={160}/>
          <DropDown options={brands} value={filters.marque}
            onChange={v=>setFilters(f=>({...f,marque:v}))}
            placeholder="Toutes les marques" minWidth={170}/>
          <DropDown options={visibleVendors} value={filters.vendeur}
            onChange={v=>setFilters(f=>({...f,vendeur:v}))}
            placeholder="Tous les conseillers" minWidth={160}/>
          {hasFilters && (
            <button onClick={()=>setFilters({etat:"",marque:"",vendeur:""})} style={{
              fontSize:11,fontFamily:"'DM Sans',sans-serif",color:"#9B2020",
              background:"none",border:"1px solid #FCEAEA",borderRadius:7,
              padding:"5px 12px",cursor:"pointer",letterSpacing:"0.04em",
            }}>Réinitialiser</button>
          )}
        </div>

        {/* Active filter tags */}
        {hasFilters && (
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>
            {filters.etat && <FilterTag label={`Statut : ${ETATS.find(s=>s.val===filters.etat)?.label}`} onRemove={()=>setFilters(f=>({...f,etat:""}))}/>}
            {filters.marque && <FilterTag label={`Marque : ${filters.marque}`} onRemove={()=>setFilters(f=>({...f,marque:""}))}/>}
            {filters.vendeur && <FilterTag label={`Conseiller : ${filters.vendeur}`} onRemove={()=>setFilters(f=>({...f,vendeur:""}))}/>}
          </div>
        )}

        {/* ── SELECT MODE TOGGLE (top-left of table) ─────────────────── */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,flexWrap:"wrap",gap:8}}>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <button onClick={toggleSelectMode} style={{
              background: bulkMode ? "#1C1510" : "#FFF",
              color: bulkMode ? "#E8DED0" : "#5A4030",
              border: bulkMode ? "1px solid #1C1510" : "1px solid #DDD4C8",
              borderRadius:8,padding:"6px 14px",fontSize:11,cursor:"pointer",
              fontFamily:"'DM Sans',sans-serif",letterSpacing:"0.06em",
              display:"flex",alignItems:"center",gap:6,fontWeight:500,
            }}>
              <span style={{
                width:14,height:14,borderRadius:3,
                border:`1.5px solid ${bulkMode?"#E8DED0":"#8A7A6A"}`,
                background: bulkMode ? "#C8A96E" : "transparent",
                display:"inline-flex",alignItems:"center",justifyContent:"center",
                fontSize:10,color:"#1C1510",fontWeight:700,
              }}>{bulkMode && "☑"}</span>
              {bulkMode ? "Mode sélection actif" : "Sélectionner"}
            </button>
            {bulkMode && (
              <button onClick={toggleSelectAllRows} style={{
                fontSize:11,color:"#5A4030",background:"#F5F0E8",
                border:"1px solid #DDD4C8",borderRadius:7,padding:"6px 12px",
                cursor:"pointer",fontFamily:"'DM Sans',sans-serif",
              }}>
                {allSelected ? "Tout décocher" : `Tout cocher (${filtered.length})`}
              </button>
            )}
          </div>
          {bulkMode && (
            <button onClick={exitBulkMode} style={{
              fontSize:11,color:"#9B2020",background:"none",
              border:"1px solid #F0D0D0",borderRadius:7,padding:"5px 12px",
              cursor:"pointer",fontFamily:"'DM Sans',sans-serif",
            }}>✕ Quitter le mode sélection</button>
          )}
        </div>

        {/* ── TABLE ───────────────────────────────────────────────────── */}
        <div style={{background:"#FFF",border:"1px solid #EDE4D5",borderRadius:14,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",minWidth: bulkMode?960:900}}>
              <thead>
                <tr>
                  {bulkMode && (
                    <th style={{...TH_STYLE,width:44,padding:"10px 8px",textAlign:"center"}}>
                      <input type="checkbox" checked={allSelected}
                        onChange={() => allSelected ? setSelected(new Set()) : setSelected(new Set(filtered.map(r=>r.id)))}
                        style={{cursor:"pointer",accentColor:"#C8A96E"}}/>
                    </th>
                  )}
                  {[
                    {col:"date",label:"Date"},
                    {col:"nom",label:"Client"},
                    {col:"marque",label:"Marque"},
                    {col:"modele",label:"Modèle"},
                    {col:"couleur",label:"Couleur"},
                    {col:"taille",label:"Taille"},
                    {col:"etat",label:"Statut"},
                    {col:"vendeur",label:"Conseiller"},
                    {col:null,label:"Actions"},
                  ].map(h => (
                    <th key={h.col||"act"} style={TH_STYLE}
                      onClick={h.col?()=>toggleSort(h.col):undefined}>
                      <span style={{display:"inline-flex",alignItems:"center",gap:4}}>
                        {h.label}
                        {h.col && sortCol===h.col && <span style={{fontSize:10}}>{sortDir==="asc"?"▲":"▼"}</span>}
                      </span>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 && (
                  <tr><td colSpan={bulkMode?10:9} style={{...TD_STYLE,textAlign:"center",color:"#B0A090",padding:"48px"}}>
                    Aucune commande trouvée
                  </td></tr>
                )}
                {filtered.map((row, i) => {
                  const expanded = expandedId === row.id;
                  const isSelected = selected.has(row.id);
                  return [
                    <tr key={row.id} className={"tr-hover " + (isSelected?"row-selected":"")}
                      style={{background:isSelected?"#FFF4E5":(i%2===0?"#FFF":"#FDFAF6")}}
                      onClick={()=>!bulkMode && setExpandedId(expanded?null:row.id)}>
                      {bulkMode && (
                        <td style={{...TD_STYLE,textAlign:"center",padding:"9px 8px"}} onClick={e=>e.stopPropagation()}>
                          <input type="checkbox" checked={isSelected}
                            onChange={()=>toggleSelect(row.id)}
                            style={{cursor:"pointer",accentColor:"#C8A96E"}}/>
                        </td>
                      )}
                      <td style={{...TD_STYLE,color:"#8A7A6A",fontSize:11,whiteSpace:"nowrap"}}>{fmtDateShortFr(row.date)}</td>
                      <td style={{...TD_STYLE,fontWeight:600,minWidth:140}}>
                        {row.nom?.toUpperCase()} <span style={{fontWeight:300,color:"#6A5A4A"}}>{row.prenom}</span>
                        {row.tel && (
                          <div onClick={(e)=>{e.stopPropagation(); setContactMenu({row, x:e.clientX, y:e.clientY});}}
                            title="Contacter le client (WhatsApp / SMS / e-mail)"
                            style={{
                              fontSize:10,color:"#1B5E9B",cursor:"pointer",
                              display:"inline-flex",alignItems:"center",gap:4,marginTop:2,
                              textDecoration:"underline",textDecorationStyle:"dotted",
                            }}>
                            📱 {fmtPhone(row.tel)}
                          </div>
                        )}
                      </td>
                      <td style={TD_STYLE}><span style={{fontWeight:600,color:"#4A2C1A"}}>{row.marque}</span></td>
                      <td style={{...TD_STYLE,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.modele}</td>
                      <td style={{...TD_STYLE,color:"#6A5A4A"}}>{(row.couleur||"").trim()}</td>
                      <td style={{...TD_STYLE,color:"#6A5A4A",whiteSpace:"nowrap"}}>{row.taille}</td>
                      <td style={TD_STYLE} onClick={e=>e.stopPropagation()}>
                        <button onClick={(e)=>setStatusMenu({row, x:e.clientX, y:e.clientY})}
                          title="Changer le statut"
                          style={{
                            all:"unset",cursor:"pointer",display:"inline-block",
                            padding:0,borderRadius:20,
                          }}>
                          {badgeEtat(row.etat, ETATS) || (
                            <span style={{
                              fontSize:10,color:"#8A7A6A",fontStyle:"italic",
                              border:"1px dashed #D8CCBE",borderRadius:20,padding:"2px 10px",
                            }}>+ statut</span>
                          )}
                        </button>
                      </td>
                      <td style={{...TD_STYLE,color:"#8A6A5A",fontSize:11,fontWeight:500}}>{row.vendeur}</td>
                      <td style={{...TD_STYLE,whiteSpace:"nowrap",padding:"8px 10px"}} onClick={e=>e.stopPropagation()}>
                        <div style={{display:"inline-flex",gap:6,alignItems:"center"}}>
                          {row.marque && (
                            <button onClick={()=>openOrderUrl(row)}
                              title={brandUrls[row.marque]?.url
                                ? `Commander sur ${brandUrls[row.marque].url}`
                                : (brandUrls[row.marque]?.email
                                    ? `Envoyer un mail à ${brandUrls[row.marque].email}`
                                    : "Renseigner URL ou e-mail dans ⚙ Paramètres")}
                              style={{
                                fontSize:12,padding:"6px 12px",
                                background: (brandUrls[row.marque]?.url || brandUrls[row.marque]?.email) ? "#1C1510" : "#F0EBE3",
                                border: (brandUrls[row.marque]?.url || brandUrls[row.marque]?.email) ? "1px solid #1C1510" : "1px solid #DDD4C8",
                                borderRadius:8,cursor:"pointer",
                                color: (brandUrls[row.marque]?.url || brandUrls[row.marque]?.email) ? "#E8DED0" : "#9A8A7A",
                                fontFamily:"'DM Sans',sans-serif",fontWeight:500,
                                display:"inline-flex",alignItems:"center",gap:5,
                              }}>🛒 <span>Commander</span></button>
                          )}
                          {row.etat === "Client prévenu" && (
                            <button onClick={()=>setTicketRow({
                              ...row,
                              articleCote: row.articleCote || [row.marque,row.modele,row.couleur&&"Couleur : "+row.couleur,row.taille&&"Taille : "+row.taille].filter(Boolean).join("\n"),
                              deCoteJusquau: row.deCoteJusquau || fmtDate(addDays(new Date(),7)),
                            })} title="Voir / imprimer le ticket caisse" style={{
                              fontSize:16,padding:"5px 10px",background:"#E3F0FC",border:"1px solid #B0CDE6",
                              borderRadius:8,cursor:"pointer",color:"#1B5E9B",lineHeight:1,
                            }}>🎫</button>
                          )}
                          <button onClick={()=>openEdit(row)} title="Modifier la fiche" style={{
                            fontSize:15,padding:"5px 10px",background:"#F5F0E8",border:"1px solid #DDD4C8",
                            borderRadius:8,cursor:"pointer",color:"#5A4030",lineHeight:1,
                          }}>✎</button>
                          <button onClick={()=>deleteRow(row.id)} title="Supprimer" style={{
                            fontSize:15,padding:"5px 10px",background:"#FDF0F0",border:"1px solid #F0D0D0",
                            borderRadius:8,cursor:"pointer",color:"#9B2020",lineHeight:1,fontWeight:600,
                          }}>✕</button>
                        </div>
                      </td>
                    </tr>,
                    expanded && (row.commentaire || row.refInt || row.refFourn || row.dateValid || row.articleCote) && (
                      <tr key={`exp-${row.id}`} style={{background:"#FFFBF5"}}>
                        <td colSpan={bulkMode?10:9} style={{padding:"8px 16px 12px 48px",borderBottom:"1px solid #F0EAE0"}}>
                          <div style={{display:"flex",gap:24,flexWrap:"wrap",fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"#6A5A4A"}}>
                            {row.refInt && <span><b>Réf. interne :</b> {row.refInt}</span>}
                            {row.refFourn && <span><b>Réf. fourn. :</b> {row.refFourn}</span>}
                            {row.dateValid && <span><b>Date validation :</b> {fmtDateFr(row.dateValid)}</span>}
                            {row.deCoteJusquau && <span><b>De côté jusqu'au :</b> {fmtDateFr(row.deCoteJusquau)}</span>}
                            {row.articleCote && <span style={{flex:"1 1 100%",whiteSpace:"pre-line"}}><b>📦 De côté :</b> {row.articleCote}</span>}
                            {row.commentaire && <span style={{flex:"1 1 100%"}}><b>💬</b> {row.commentaire}</span>}
                          </div>
                        </td>
                      </tr>
                    ),
                  ].filter(Boolean);
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* ── BULK ACTION BAR (sticky bottom) ──────────────────────────── */}
      {bulkMode && selected.size > 0 && (
        <div style={{
          position:"fixed",bottom:0,left:0,right:0,zIndex:80,
          background:"rgba(28,21,16,0.97)",backdropFilter:"blur(10px)",
          padding:"14px 24px",
          display:"flex",alignItems:"center",justifyContent:"space-between",
          flexWrap:"wrap",gap:12,boxShadow:"0 -8px 24px rgba(0,0,0,0.15)",
        }}>
          <div style={{display:"flex",alignItems:"center",gap:12,color:"#F5EDE0",fontFamily:"'DM Sans',sans-serif"}}>
            <span style={{
              background:"#C8A96E",color:"#1C1510",borderRadius:20,
              padding:"3px 10px",fontSize:12,fontWeight:700,
            }}>{selected.size}</span>
            <span style={{fontSize:13}}>commande{selected.size>1?"s":""} sélectionnée{selected.size>1?"s":""}</span>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span style={{color:"#A09080",fontSize:11,fontFamily:"'DM Sans',sans-serif",letterSpacing:"0.06em",textTransform:"uppercase"}}>
              Changer le statut →
            </span>
            <select onChange={e => e.target.value && applyBulkStatus(e.target.value)}
              defaultValue="" style={{
                fontSize:13,fontFamily:"'DM Sans',sans-serif",
                background:"#F5EDE0",color:"#1C1510",
                border:"none",borderRadius:8,padding:"7px 14px",
                cursor:"pointer",minWidth:200,fontWeight:500,
              }}>
              <option value="" disabled>Choisir un nouveau statut…</option>
              {ETATS.map(s => <option key={s.val} value={s.val}>{s.label}</option>)}
            </select>
            <button onClick={exitBulkMode} style={{
              background:"transparent",color:"#C8A96E",border:"1px solid #C8A96E",
              borderRadius:8,padding:"7px 14px",fontSize:12,cursor:"pointer",
              fontFamily:"'DM Sans',sans-serif",
            }}>Annuler</button>
          </div>
        </div>
      )}

      {/* ── EDIT MODAL ─────────────────────────────────────────────── */}
      {showEdit && (
        <Modal onClose={()=>setShowEdit(false)} title={editId ? "Modifier la commande" : "Nouvelle commande"}>
          <div style={{padding:"22px 28px"}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <FormField label="Date création">
                <input type="date" value={form.date||""} onChange={e=>setForm(p=>({...p,date:e.target.value}))} style={inputStyle()}/>
              </FormField>
              <FormField label="Statut" full>
                <select value={form.etat||""} onChange={e=>setForm(p=>({...p,etat:e.target.value}))} style={inputStyle()}>
                  <option value="">— Sélectionner —</option>
                  {ETATS.map(s => <option key={s.val} value={s.val}>{s.label}</option>)}
                </select>
              </FormField>
              <FormField label="Nom client *">
                <input
                  required
                  value={form.nom||""}
                  onChange={e=>setForm(p=>({...p,nom: upperCaseFr(e.target.value)}))}
                  style={{...inputStyle(),fontWeight:600}}
                  placeholder="DUPONT"/>
              </FormField>
              <FormField label="Prénom">
                <input
                  value={form.prenom||""}
                  onChange={e=>setForm(p=>({...p,prenom: nameCase(e.target.value)}))}
                  style={inputStyle()}
                  placeholder="Marie"/>
              </FormField>
              <FormField label="Téléphone *">
                <input
                  required
                  value={form.tel||""}
                  onChange={e=>setForm(p=>({...p,tel:e.target.value}))}
                  style={inputStyle()} placeholder="06 12 34 56 78"/>
              </FormField>
              <FormField label="E-mail">
                <input type="email" value={form.email||""}
                  onChange={e=>setForm(p=>({...p,email:e.target.value.toLowerCase()}))}
                  style={inputStyle()} placeholder="client@email.com"/>
              </FormField>
              <FormField label="Conseiller">
                <select value={form.vendeur||""} onChange={e=>setForm(p=>({...p,vendeur:e.target.value}))} style={inputStyle()}>
                  <option value="">—</option>
                  {visibleVendors.map(v => <option key={v} value={v}>{v}</option>)}
                </select>
              </FormField>
              <FormField label="Marque">
                <select value={form.marque||""} onChange={e=>setForm(p=>({...p,marque:e.target.value}))} style={inputStyle()}>
                  <option value="">—</option>
                  {brands.map(b => <option key={b} value={b}>{b}</option>)}
                  {form.marque && !brands.includes(form.marque) && <option value={form.marque}>{form.marque} (hors liste)</option>}
                </select>
              </FormField>
              <FormField label="Nom modèle">
                <input value={form.modele||""}
                  onChange={e=>setForm(p=>({...p,modele: upperCaseFr(e.target.value)}))}
                  style={inputStyle()}/>
              </FormField>
              <FormField label="Réf. interne">
                <input value={form.refInt||""}
                  onChange={e=>setForm(p=>({...p,refInt: upperCaseFr(e.target.value)}))}
                  style={inputStyle()} placeholder="Saisie libre"/>
              </FormField>
              <FormField label="Réf. fournisseur">
                <input value={form.refFourn||""}
                  onChange={e=>setForm(p=>({...p,refFourn: upperCaseFr(e.target.value)}))}
                  style={inputStyle()} placeholder="Saisie libre"/>
              </FormField>
              <FormField label="Couleur">
                {(() => {
                  const isOther = form.couleur && !COLORS_TOP15.includes(form.couleur.trim());
                  return (
                    <>
                      <select
                        value={isOther ? "__autre__" : (form.couleur || "")}
                        onChange={e => {
                          const v = e.target.value;
                          if (v === "__autre__") setForm(p => ({...p, couleur: " "}));
                          else setForm(p => ({...p, couleur: v}));
                        }}
                        style={inputStyle()}>
                        <option value="">—</option>
                        {COLORS_TOP15.map(c => <option key={c} value={c}>{c}</option>)}
                        <option value="__autre__">Autre…</option>
                      </select>
                      {isOther && (
                        <input
                          autoFocus
                          value={form.couleur.trim()}
                          onChange={e => setForm(p => ({...p, couleur: upperCaseFr(e.target.value) || " "}))}
                          placeholder="PRÉCISER LA COULEUR"
                          style={{...inputStyle(), marginTop:6}}/>
                      )}
                    </>
                  );
                })()}
              </FormField>
              <FormField label="Taille">
                <input value={form.taille||""}
                  onChange={e=>setForm(p=>({...p,taille: upperCaseFr(e.target.value)}))}
                  style={inputStyle()}/>
              </FormField>
            </div>
            <FormField label="Commentaire">
              <textarea value={form.commentaire||""}
                onChange={e=>setForm(p=>({...p,commentaire: upperCaseFr(e.target.value)}))}
                rows={3} style={{...inputStyle(),resize:"vertical",minHeight:60}}/>
            </FormField>
            <div style={{
              marginTop:10,padding:"8px 12px",
              background:"#F8F3EB",borderRadius:8,
              fontSize:10,color:"#7A5A2A",fontFamily:"'DM Sans',sans-serif",
            }}>
              ℹ️ La date de validation sera automatiquement mise à jour à la date du jour lors d'un changement de statut.
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button onClick={()=>setShowEdit(false)} style={{
                flex:1,padding:"11px",background:"#F5F0E8",border:"1px solid #DDD4C8",borderRadius:9,
                fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",color:"#5A4030",
              }}>Annuler</button>
              <button onClick={saveForm} style={{
                flex:2,padding:"11px",background:"#1C1510",border:"none",borderRadius:9,
                fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",
                color:"#E8DED0",letterSpacing:"0.06em",
              }}>{editId ? "Enregistrer" : "Ajouter la commande"}</button>
            </div>
          </div>
        </Modal>
      )}

      {/* ── TICKET CAISSE MODAL ─────────────────────────────────────── */}
      {ticketRow && (
        <Modal onClose={()=>{saveTicket(); setTicketRow(null);}} title="Ticket caisse — Client prévenu" maxWidth={560}>
          <div style={{padding:"20px 24px"}}>
            <div style={{
              display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14,
            }}>
              <FormField label="De côté jusqu'au">
                <input type="date" value={ticketRow.deCoteJusquau||""}
                  onChange={e=>setTicketRow(r=>({...r,deCoteJusquau:e.target.value}))}
                  style={inputStyle()}/>
              </FormField>
              <FormField label="Conseiller">
                <select value={ticketRow.vendeur||""}
                  onChange={e=>setTicketRow(r=>({...r,vendeur:e.target.value}))}
                  style={inputStyle()}>
                  <option value="">—</option>
                  {visibleVendors.map(v => <option key={v} value={v}>{v}</option>)}
                </select>
              </FormField>
            </div>

            {/* 3 cases à cocher contact client (sauvegardées avec la commande) */}
            <div style={{
              background:"#F0F5FB",border:"1px solid #D6E4F0",borderRadius:10,
              padding:"12px 14px",marginBottom:18,
            }}>
              <div style={{
                fontSize:10,letterSpacing:"0.12em",color:"#1B5E9B",
                fontWeight:700,marginBottom:8,textTransform:"uppercase",
              }}>Contact client</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:14}}>
                {[
                  {key:"msgVocal",    label:"📞 Message vocal",   color:"#5C7AA0"},
                  {key:"msgWhatsapp", label:"💬 Message WhatsApp", color:"#1FA855"},
                  {key:"mailEnvoye",  label:"✉️ Mail envoyé",     color:"#A0620A"},
                ].map(c => (
                  <label key={c.key} style={{
                    display:"flex",alignItems:"center",gap:8,cursor:"pointer",
                    fontFamily:"'DM Sans',sans-serif",fontSize:12,color:"#1C1510",
                    background:"#FFF",border:"1px solid #D6E4F0",borderRadius:8,
                    padding:"6px 12px",
                  }}>
                    <input type="checkbox" checked={!!ticketRow[c.key]}
                      onChange={e=>{
                        const v = e.target.checked;
                        setTicketRow(r=>({...r, [c.key]: v}));
                        updateRow(ticketRow.id, { [c.key]: v });
                      }}
                      style={{width:16,height:16,accentColor:c.color,cursor:"pointer"}}/>
                    {c.label}
                  </label>
                ))}
              </div>
              {!ticketRow.deCoteJusquau && ticketRow.msgVocal && (
                <div style={{marginTop:10,fontSize:11,color:"#5C7AA0",fontFamily:"'DM Sans',sans-serif"}}>
                  Aucune date « de côté » → le ticket imprimera <b>Message vocal du {fmtDateFr(fmtDate(new Date()))}</b>
                  et indiquera la relance au <b>{fmtDateFr(fmtDate(addDays(new Date(),5)))}</b>.
                </div>
              )}
            </div>
            <FormField label="Quel article reste de côté ?">
              <textarea value={ticketRow.articleCote||""}
                onChange={e=>setTicketRow(r=>({...r,articleCote: upperCaseFr(e.target.value)}))}
                rows={4} style={{...inputStyle(),resize:"vertical",minHeight:80,fontFamily:"'Courier New',monospace",fontWeight:600}}/>
            </FormField>

            <div style={{
              marginTop:18,marginBottom:18,
              background:"#F8F3EB",padding:"14px",borderRadius:10,
              fontSize:10,color:"#8A7A6A",fontFamily:"'DM Sans',sans-serif",
              textAlign:"center",letterSpacing:"0.06em",textTransform:"uppercase",
            }}>↓ Aperçu du ticket caisse ↓</div>

            <div className="ticket-printable">
              <div className="ticket-paper">
                {/* Nom du magasin en grand et gras (à la place du logo) */}
                <div style={{textAlign:"center",paddingBottom:6,paddingTop:2}}>
                  <div style={{
                    fontSize:18,fontWeight:900,letterSpacing:"0.12em",
                    color:"#1C1510",lineHeight:1.1,
                  }}>{magasinNom(ticketRow.magasin || activeMagasin || "SQUARE")}</div>
                </div>
                <div style={{textAlign:"center",fontSize:12,letterSpacing:"0.16em",borderTop:"2px dashed #1C1510",borderBottom:"2px dashed #1C1510",padding:"6px 0",margin:"6px 0 12px"}}>
                  <b style={{fontWeight:900,fontSize:13}}>COMMANDE CLIENT</b>
                </div>

                {/* Identité client — gros et gras */}
                <div style={{textAlign:"center",fontSize:15,lineHeight:1.4,fontWeight:900,marginBottom:4}}>
                  {(ticketRow.nom||"").toUpperCase()} {ticketRow.prenom||""}
                </div>
                <div style={{textAlign:"center",fontSize:13,fontWeight:800,marginBottom:12}}>
                  {ticketRow.tel ? fmtPhone(ticketRow.tel) : "—"}
                </div>

                {/* Bloc DE CÔTÉ JUSQU'AU ou MESSAGE VOCAL */}
                {ticketRow.deCoteJusquau ? (
                  <div style={{margin:"12px 0",borderTop:"1px dashed #1C1510",borderBottom:"1px dashed #1C1510",padding:"10px 0",textAlign:"center"}}>
                    <div style={{fontSize:10,marginBottom:4,letterSpacing:"0.12em",fontWeight:800}}>DE CÔTÉ JUSQU'AU</div>
                    <div style={{fontSize:19,fontWeight:900,letterSpacing:"0.06em"}}>
                      {fmtDateFr(ticketRow.deCoteJusquau)}
                    </div>
                  </div>
                ) : ticketRow.msgVocal ? (
                  <div style={{margin:"12px 0",borderTop:"1px dashed #1C1510",borderBottom:"1px dashed #1C1510",padding:"10px 0",textAlign:"center"}}>
                    <div style={{fontSize:10,marginBottom:4,letterSpacing:"0.12em",fontWeight:800}}>MESSAGE VOCAL CE JOUR</div>
                    <div style={{fontSize:16,fontWeight:900,letterSpacing:"0.06em"}}>
                      {fmtDateFr(fmtDate(new Date()))}
                    </div>
                  </div>
                ) : null}

                {/* Conseiller */}
                <div style={{textAlign:"center",margin:"10px 0",fontSize:11}}>
                  <div style={{fontSize:9,letterSpacing:"0.12em",marginBottom:2,fontWeight:800}}>CONSEILLER</div>
                  <b style={{fontSize:13,fontWeight:900}}>{ticketRow.vendeur||"—"}</b>
                </div>

                {/* Article réservé */}
                {ticketRow.articleCote && (
                  <div style={{marginTop:12,paddingTop:8,borderTop:"1px dashed #1C1510",fontSize:12,whiteSpace:"pre-line",textAlign:"center",fontWeight:800,lineHeight:1.5}}>
                    <div style={{fontSize:9,letterSpacing:"0.12em",marginBottom:4,fontWeight:800}}>ARTICLE RÉSERVÉ</div>
                    {ticketRow.articleCote}
                  </div>
                )}

                {/* Pied : REMETTRE EN RAYON LE… ou RELANCER LE CLIENT À PARTIR DU… */}
                {ticketRow.deCoteJusquau ? (
                  <div style={{textAlign:"center",marginTop:14,paddingTop:8,borderTop:"2px dashed #1C1510",fontSize:11}}>
                    <div style={{fontSize:9,letterSpacing:"0.12em",marginBottom:3,fontWeight:800}}>REMETTRE EN RAYON LE</div>
                    <div style={{fontSize:15,fontWeight:900,letterSpacing:"0.06em"}}>
                      {fmtDateFr(fmtDate(addDays(parseDate(ticketRow.deCoteJusquau) || new Date(), 7)))}
                    </div>
                  </div>
                ) : ticketRow.msgVocal ? (
                  <div style={{textAlign:"center",marginTop:14,paddingTop:8,borderTop:"2px dashed #1C1510",fontSize:11}}>
                    <div style={{fontSize:9,letterSpacing:"0.12em",marginBottom:3,fontWeight:800}}>RELANCER LE CLIENT À PARTIR DU</div>
                    <div style={{fontSize:15,fontWeight:900,letterSpacing:"0.06em"}}>
                      {fmtDateFr(fmtDate(addDays(new Date(), 5)))}
                    </div>
                  </div>
                ) : null}
              </div>
            </div>

            <div style={{display:"flex",gap:10,marginTop:22}}>
              <button onClick={()=>{saveTicket(); setTicketRow(null);}} style={{
                flex:1,padding:"11px",background:"#F5F0E8",border:"1px solid #DDD4C8",borderRadius:9,
                fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",color:"#5A4030",
              }}>Fermer</button>
              <button onClick={()=>{saveTicket(); printTicket();}} style={{
                flex:2,padding:"11px",background:"#1C1510",border:"none",borderRadius:9,
                fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",
                color:"#E8DED0",letterSpacing:"0.06em",display:"flex",alignItems:"center",justifyContent:"center",gap:8,
              }}>🖨 Imprimer le ticket</button>
            </div>
          </div>
        </Modal>
      )}

      {/* ── SETTINGS MODAL (manage brands / vendors) ─────────────────── */}
      {showSettings && (
        <Modal onClose={()=>setShowSettings(false)} title="Gérer les listes" maxWidth={520}>
          <div style={{padding:"0 0 20px"}}>
            <div style={{display:"flex",borderBottom:"1px solid #EDE4D5",padding:"0 24px"}}>
              {[
                {key:"brands",label:`Fournisseurs (${brands.length})`},
                {key:"vendors",label:`Conseillers (${vendeursAll.length})`},
                {key:"statuts",label:`Statuts (${statutsAll.length})`},
              ].map(t => (
                <button key={t.key} onClick={()=>setSettingsTab(t.key)} style={{
                  all:"unset",cursor:"pointer",padding:"12px 18px",
                  fontFamily:"'DM Sans',sans-serif",fontSize:12,
                  color: settingsTab===t.key?"#1C1510":"#A09080",
                  borderBottom: settingsTab===t.key?"2px solid #C8A96E":"2px solid transparent",
                  letterSpacing:"0.06em",fontWeight:500,
                }}>{t.label}</button>
              ))}
            </div>

            <div style={{padding:"18px 24px",maxHeight:380,overflowY:"auto"}}>
              {settingsTab === "brands" ? (
                <>
                  <div style={{display:"flex",gap:8,marginBottom:12}}>
                    <input value={newBrandInput} onChange={e=>setNewBrandInput(e.target.value)}
                      onKeyDown={e=>e.key==="Enter"&&addBrand()}
                      placeholder="Nouveau fournisseur (sera mis en majuscules)"
                      style={inputStyle()}/>
                    <button onClick={addBrand} style={{
                      background:"#1C1510",color:"#E8DED0",border:"none",borderRadius:8,
                      padding:"0 16px",fontSize:12,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",whiteSpace:"nowrap",
                    }}>+ Ajouter</button>
                  </div>
                  <div style={{fontSize:10,color:"#A09080",fontFamily:"'DM Sans',sans-serif",marginBottom:10,letterSpacing:"0.04em"}}>
                    URL B2B, téléphone et e-mail du contact commercial. L'URL est utilisée par le bouton 🛒 Commander.
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:10}}>
                    {brands.map(b => {
                      const info = brandUrls[b] || {url:"",tel:"",email:""};
                      const fieldStyle = {
                        flex:1,padding:"6px 9px",fontSize:11,
                        fontFamily:"'DM Sans',sans-serif",
                        border:"1px solid #E0D8CE",borderRadius:7,
                        background:"#FFF",color:"#1C1510",minWidth:0,
                      };
                      return (
                        <div key={b} style={{
                          background:"#FAFAF8",border:"1px solid #EDE4D5",borderRadius:10,
                          padding:"10px 12px",
                        }}>
                          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:7}}>
                            <span style={{
                              fontSize:12,fontFamily:"'DM Sans',sans-serif",fontWeight:700,
                              color:"#4A2C1A",letterSpacing:"0.04em",
                            }}>{b}</span>
                            <div style={{display:"flex",gap:6}}>
                              {info.url && (
                                <button onClick={()=>openOrderUrl(b)} title="Tester le lien" style={{
                                  background:"#E3F0FC",border:"1px solid #B0CDE6",borderRadius:6,
                                  padding:"4px 8px",fontSize:11,cursor:"pointer",color:"#1B5E9B",
                                }}>↗ Ouvrir</button>
                              )}
                              <button onClick={()=>removeBrand(b)} title="Supprimer" style={{
                                background:"#FDEEEE",border:"none",borderRadius:"50%",
                                width:24,height:24,cursor:"pointer",color:"#9B2020",fontSize:14,lineHeight:1,
                              }}>×</button>
                            </div>
                          </div>
                          <div style={{display:"grid",gridTemplateColumns:"1fr",gap:6}}>
                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                              <span style={{fontSize:11,width:18,textAlign:"center"}}>🔗</span>
                              <input value={info.url||""} onChange={e=>setBrandField(b,"url",e.target.value)}
                                placeholder="URL B2B (ex : b2b.marque.com)" style={fieldStyle}/>
                            </div>
                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                              <span style={{fontSize:11,width:18,textAlign:"center"}}>📞</span>
                              <input value={info.tel||""} onChange={e=>setBrandField(b,"tel",e.target.value)}
                                placeholder="Téléphone du contact commercial" style={fieldStyle}/>
                            </div>
                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                              <span style={{fontSize:11,width:18,textAlign:"center"}}>✉️</span>
                              <input type="email" value={info.email||""} onChange={e=>setBrandField(b,"email",e.target.value)}
                                placeholder="E-mail du contact" style={fieldStyle}/>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </>
              ) : settingsTab === "vendors" ? (
                <>
                  <div style={{display:"flex",gap:8,marginBottom:12}}>
                    <input value={newVendorInput} onChange={e=>setNewVendorInput(e.target.value)}
                      onKeyDown={e=>e.key==="Enter"&&addVendor()}
                      placeholder="Nouveau conseiller (sera mis en majuscules)"
                      style={inputStyle()}/>
                    <button onClick={addVendor} disabled={!isManager} style={{
                      background: isManager ? "#1C1510" : "#AAA",
                      color:"#E8DED0",border:"none",borderRadius:8,
                      padding:"0 16px",fontSize:12,
                      cursor: isManager ? "pointer" : "not-allowed",
                      fontFamily:"'DM Sans',sans-serif",whiteSpace:"nowrap",
                    }}>+ Ajouter</button>
                  </div>
                  {!isManager && (
                    <div style={{
                      background:"#FFF8EB",border:"1px solid #E8DCC0",borderRadius:8,
                      padding:"8px 12px",fontSize:11,color:"#7A5A2A",marginBottom:12,
                      fontFamily:"'DM Sans',sans-serif",
                    }}>ℹ️ Seuls les managers peuvent modifier la liste des conseillers.</div>
                  )}
                  <div style={{display:"flex",flexDirection:"column",gap:8}}>
                    {vendeursAll.map(v => (
                      <div key={v.id} style={{
                        background:"#FAFAF8",border:"1px solid #EDE4D5",borderRadius:10,
                        padding:"10px 12px",
                      }}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
                          <span style={{
                            fontSize:13,fontFamily:"'DM Sans',sans-serif",fontWeight:700,
                            color:"#4A2C1A",letterSpacing:"0.04em",
                          }}>{v.nom}</span>
                          {isManager && (
                            <button onClick={()=>removeVendeur(v.id, v.nom)} title="Supprimer" style={{
                              background:"#FDEEEE",border:"none",borderRadius:"50%",
                              width:24,height:24,cursor:"pointer",color:"#9B2020",fontSize:14,lineHeight:1,
                            }}>×</button>
                          )}
                        </div>
                        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                          {MAGASIN_CODES.map(code => {
                            const checked = v.magasins.includes(code);
                            return (
                              <label key={code} style={{
                                display:"flex",alignItems:"center",gap:5,
                                background: checked ? "#FFF8EB" : "#FFF",
                                border: `1px solid ${checked ? "#C8A96E" : "#E0D8CE"}`,
                                borderRadius:8,padding:"4px 9px",
                                fontSize:11,fontFamily:"'DM Sans',sans-serif",
                                color: checked ? "#7A5A2A" : "#8A7A6A",
                                cursor: isManager ? "pointer" : "default",
                                fontWeight: checked ? 700 : 500,
                                opacity: isManager ? 1 : 0.7,
                              }}>
                                <input type="checkbox" checked={checked}
                                  disabled={!isManager}
                                  onChange={()=>toggleVendeurMagasin(v, code)}
                                  style={{accentColor:"#C8A96E",cursor:isManager?"pointer":"default"}}/>
                                {MAGASINS[code]?.nom?.replace("CLOANE ", "") || code}
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                  </div>
                </>
              ) : (
                /* ── ONGLET STATUTS ── */
                <>
                  {!isManager ? (
                    <div style={{
                      background:"#FFF8EB",border:"1px solid #E8DCC0",borderRadius:8,
                      padding:"12px 14px",fontSize:11,color:"#7A5A2A",
                      fontFamily:"'DM Sans',sans-serif",lineHeight:1.5,
                    }}>ℹ️ Seuls les managers peuvent ajouter, modifier ou supprimer des statuts.</div>
                  ) : null}

                  {isManager && (
                    <div style={{
                      background:"#F0F5FB",border:"1px solid #D6E4F0",borderRadius:8,
                      padding:"10px 12px",fontSize:11,color:"#1B5E9B",marginBottom:14,
                      fontFamily:"'DM Sans',sans-serif",lineHeight:1.5,
                    }}>
                      💡 Cliquez sur un nom ou une couleur pour modifier.
                      Le renommage met à jour <b>toutes</b> les commandes utilisant ce statut.
                    </div>
                  )}

                  <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:14}}>
                    {statutsAll.map((s, idx) => (
                      <div key={s.id} style={{
                        background:"#FAFAF8",border:"1px solid #EDE4D5",borderRadius:10,
                        padding:"10px 12px",display:"flex",alignItems:"center",gap:10,
                      }}>
                        {/* Couleur principale + aperçu pastille */}
                        <div style={{display:"flex",alignItems:"center",gap:8,minWidth:0,flex:1}}>
                          <input type="color" value={s.couleur || "#8A7A6A"}
                            disabled={!isManager}
                            onChange={e=>updateStatutColor(s.id, e.target.value)}
                            style={{
                              width:30,height:30,border:"1px solid #E0D8CE",borderRadius:6,
                              cursor: isManager ? "pointer" : "default",background:"#FFF",padding:2,
                            }}/>
                          <span style={{
                            display:"inline-block",padding:"3px 10px",borderRadius:20,
                            fontSize:10,fontWeight:600,fontFamily:"'DM Sans',sans-serif",
                            color: s.couleur, background: lightenColor(s.couleur),
                            whiteSpace:"nowrap",flexShrink:0,
                          }}>{s.nom}</span>
                          <input value={s.nom}
                            disabled={!isManager}
                            onChange={e=>renameStatutInline(s.id, s.nom, e.target.value)}
                            onBlur={e=>commitStatutRename(s.id, s.nom, e.target.value)}
                            style={{
                              flex:1,padding:"6px 9px",fontSize:12,
                              fontFamily:"'DM Sans',sans-serif",fontWeight:600,
                              border:"1px solid #E0D8CE",borderRadius:7,
                              background:"#FFF",color:"#1C1510",minWidth:80,
                            }}/>
                        </div>
                        {isManager && (
                          <button onClick={()=>removeStatut(s)} title="Supprimer ce statut" style={{
                            background:"#FDEEEE",border:"none",borderRadius:"50%",
                            width:24,height:24,cursor:"pointer",color:"#9B2020",fontSize:14,lineHeight:1,flexShrink:0,
                          }}>×</button>
                        )}
                      </div>
                    ))}
                  </div>

                  {isManager && (
                    <button onClick={addStatut} style={{
                      background:"#1C1510",color:"#E8DED0",border:"none",borderRadius:8,
                      padding:"9px 16px",fontSize:12,cursor:"pointer",
                      fontFamily:"'DM Sans',sans-serif",fontWeight:600,
                    }}>+ Ajouter un statut</button>
                  )}
                </>
              )}
            </div>
          </div>
        </Modal>
      )}

      {/* ── CONTACT CLIENT POPOVER (clic sur le téléphone) ───────────── */}
      {contactMenu && (
        <div onClick={()=>setContactMenu(null)} style={{
          position:"fixed",inset:0,zIndex:1100,background:"transparent",
        }}>
          <div onClick={e=>e.stopPropagation()} style={{
            position:"fixed",
            top: Math.min(contactMenu.y + 8, (typeof window!=="undefined"?window.innerHeight:800) - 230),
            left: Math.min(contactMenu.x, (typeof window!=="undefined"?window.innerWidth:1000) - 250),
            background:"#FFFCF8",borderRadius:14,
            border:"1px solid #EDE4D5",boxShadow:"0 16px 48px rgba(0,0,0,0.18)",
            width:240,overflow:"hidden",
          }}>
            <div style={{padding:"12px 16px",borderBottom:"1px solid #EDE4D5"}}>
              <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,color:"#1C1510"}}>
                {(contactMenu.row.nom||"").toUpperCase()} {contactMenu.row.prenom||""}
              </div>
              <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"#8A7A6A",marginTop:2}}>
                {fmtPhone(contactMenu.row.tel)}
              </div>
              <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"#A09080",marginTop:6,lineHeight:1.4}}>
                Message « commande arrivée » prérempli
              </div>
            </div>
            <button onClick={()=>sendWhatsApp(contactMenu.row)} style={contactBtnStyle("#1FA855")}>
              <span>💬</span> WhatsApp
            </button>
            <button onClick={()=>sendSMS(contactMenu.row)} style={contactBtnStyle("#5C7AA0")}>
              <span>📩</span> SMS
            </button>
            <button onClick={()=>sendMail(contactMenu.row)}
              style={{...contactBtnStyle("#A0620A"), opacity: contactMenu.row.email?1:0.5}}>
              <span>✉️</span> E-mail {!contactMenu.row.email && <span style={{fontSize:9,color:"#A09080"}}>(non renseigné)</span>}
            </button>
          </div>
        </div>
      )}

      {/* ── RUPTURE STOCK : modal pour prévenir le client ─────────────── */}
      {ruptureRow && (
        <Modal onClose={()=>setRuptureRow(null)} title="Rupture de stock — prévenir le client" maxWidth={520}>
          <div style={{
            background:"#FCEAEA",border:"1px solid #F0C8C8",borderRadius:10,
            padding:"14px 16px",marginBottom:16,
          }}>
            <div style={{fontSize:13,color:"#7A2020",fontWeight:600,marginBottom:8}}>
              ⚠️ Article en rupture — informer le client
            </div>
            <div style={{fontSize:12,color:"#5A4030",lineHeight:1.5,fontFamily:"'DM Sans',sans-serif"}}>
              Cette commande passe en <b>Rupture Stock</b>. Choisissez un moyen pour prévenir
              <b> {(ruptureRow.nom||"").toUpperCase()} {ruptureRow.prenom||""}</b> avec un message d'excuses prérempli
              l'invitant à recontacter le magasin plus tard dans la saison.
            </div>
          </div>

          {/* Aperçu du message */}
          <div style={{
            background:"#FFFCF8",border:"1px solid #EDE4D5",borderRadius:10,
            padding:"14px 16px",marginBottom:18,fontSize:12,
            fontFamily:"'DM Sans',sans-serif",color:"#1C1510",
            whiteSpace:"pre-line",lineHeight:1.55,
          }}>
            <div style={{fontSize:9,letterSpacing:"0.12em",color:"#8A7A6A",fontWeight:700,marginBottom:8,textTransform:"uppercase"}}>
              Aperçu du message
            </div>
            {buildRuptureMessage(ruptureRow, ruptureRow.magasin)}
          </div>

          {/* Boutons de contact */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10}}>
            <button onClick={()=>sendRuptureWhatsApp(ruptureRow)} style={contactBtnInlineStyle("#1FA855")}>
              💬 WhatsApp
            </button>
            <button onClick={()=>sendRuptureSMS(ruptureRow)} style={contactBtnInlineStyle("#5C7AA0")}>
              📩 SMS
            </button>
            <button onClick={()=>sendRuptureMail(ruptureRow)}
              style={{...contactBtnInlineStyle("#A0620A"), opacity: ruptureRow.email?1:0.5}}
              disabled={!ruptureRow.email}>
              ✉️ E-mail{!ruptureRow.email && " (non renseigné)"}
            </button>
          </div>

          <div style={{display:"flex",gap:10,marginTop:18,justifyContent:"flex-end"}}>
            <button onClick={()=>setRuptureRow(null)} style={{
              padding:"9px 18px",background:"#F5F0E8",border:"1px solid #DDD4C8",
              borderRadius:8,fontSize:12,cursor:"pointer",color:"#5A4030",
              fontFamily:"'DM Sans',sans-serif",
            }}>Fermer</button>
          </div>
        </Modal>
      )}

      {/* ── STATUS POPOVER (clic sur le statut d'une ligne) ───────────── */}
      {statusMenu && (
        <div onClick={()=>setStatusMenu(null)} style={{
          position:"fixed",inset:0,zIndex:1100,background:"transparent",
        }}>
          <div onClick={e=>e.stopPropagation()} style={{
            position:"fixed",
            top: Math.min(statusMenu.y + 8, (typeof window!=="undefined"?window.innerHeight:800) - 380),
            left: Math.min(statusMenu.x, (typeof window!=="undefined"?window.innerWidth:1000) - 260),
            background:"#FFFCF8",borderRadius:14,
            border:"1px solid #EDE4D5",boxShadow:"0 16px 48px rgba(0,0,0,0.18)",
            width:250,overflow:"hidden",
          }}>
            <div style={{padding:"11px 16px",borderBottom:"1px solid #EDE4D5",background:"#FAF5EC"}}>
              <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"#8A7A6A",letterSpacing:"0.12em",textTransform:"uppercase",fontWeight:600,marginBottom:3}}>
                Changer le statut
              </div>
              <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:15,color:"#1C1510"}}>
                {(statusMenu.row.nom||"").toUpperCase()} {statusMenu.row.prenom||""}
              </div>
            </div>
            {ETATS.map(s => {
              const isCurrent = statusMenu.row.etat === s.val;
              return (
                <button key={s.val}
                  onClick={()=>{
                    if (!isCurrent) {
                      const today = fmtDate(new Date());
                      updateRow(statusMenu.row.id, {etat: s.val, dateValid: today});
                      if (s.val === "Client prévenu") {
                        const r = statusMenu.row;
                        setTimeout(()=>setTicketRow({
                          ...r, etat: s.val, dateValid: today,
                          articleCote: r.articleCote || [r.marque, r.modele, r.couleur && "Couleur : "+r.couleur, r.taille && "Taille : "+r.taille].filter(Boolean).join("\n"),
                          deCoteJusquau: r.deCoteJusquau || fmtDate(addDays(new Date(), 7)),
                        }), 150);
                      }
                      if (s.val === "Rupture Stock") {
                        const r = statusMenu.row;
                        setTimeout(()=>setRuptureRow({...r, etat: s.val, dateValid: today}), 150);
                      }
                    }
                    setStatusMenu(null);
                  }}
                  style={{
                    display:"flex",alignItems:"center",gap:10,width:"100%",
                    padding:"9px 16px",background: isCurrent?"#F0EBE3":"transparent",
                    border:"none",borderBottom:"1px solid #F4EFE6",cursor:"pointer",
                    fontFamily:"'DM Sans',sans-serif",fontSize:12,
                    color: isCurrent?"#1C1510":"#2C1E0F",
                    textAlign:"left",borderLeft:`3px solid ${s.color}`,
                    fontWeight: isCurrent?700:500,
                  }}>
                  <span style={{flex:1}}>{s.label}</span>
                  {isCurrent && <span style={{fontSize:10,color:s.color}}>● actuel</span>}
                </button>
              );
            })}
          </div>
        </div>
      )}

    </div>
  );
}

function contactBtnStyle(accent) {
  return {
    display:"flex",alignItems:"center",gap:10,width:"100%",
    padding:"11px 16px",background:"transparent",border:"none",
    borderBottom:"1px solid #F4EFE6",cursor:"pointer",
    fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"#2C1E0F",
    textAlign:"left",borderLeft:`3px solid ${accent}`,
  };
}
function contactBtnInlineStyle(accent) {
  return {
    display:"flex",alignItems:"center",justifyContent:"center",gap:8,
    padding:"12px 16px",background:"#FFF",
    border:`1.5px solid ${accent}`,cursor:"pointer",
    fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:600,
    color:accent,borderRadius:9,
  };
}

// ── Inline helper components ────────────────────────────────────────────────

function Modal({children, onClose, title, maxWidth=640}) {
  return (
    <div onClick={onClose} style={{
      position:"fixed",inset:0,zIndex:1000,
      background:"rgba(28,21,16,0.55)",backdropFilter:"blur(4px)",
      display:"flex",alignItems:"flex-start",justifyContent:"center",
      padding:"32px 16px 24px",overflowY:"auto",
    }}>
      <div onClick={e=>e.stopPropagation()} style={{
        background:"#FFFCF8",borderRadius:18,width:"100%",maxWidth,
        maxHeight:"92vh",overflowY:"auto",
        boxShadow:"0 24px 80px rgba(0,0,0,0.2)",
        border:"1px solid #EDE4D5",
      }}>
        <div style={{
          padding:"20px 28px 16px",borderBottom:"1px solid #EDE4D5",
          display:"flex",justifyContent:"space-between",alignItems:"center",
          position:"sticky",top:0,background:"#FFFCF8",borderRadius:"18px 18px 0 0",zIndex:1,
        }}>
          <div style={{
            fontFamily:"'Cormorant Garamond',serif",fontSize:20,fontWeight:400,
            color:"#1C1510",letterSpacing:"0.06em",
          }}>{title}</div>
          <button onClick={onClose} style={{
            background:"none",border:"none",fontSize:18,cursor:"pointer",color:"#A09080",padding:4,
          }}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function FormField({label, children, full}) {
  return (
    <div style={full ? {gridColumn:"span 2"} : {}}>
      <label style={{
        display:"block",fontSize:10,color:"#8A7A6A",
        fontFamily:"'DM Sans',sans-serif",fontWeight:600,
        textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:5,
      }}>{label}</label>
      {children}
    </div>
  );
}

function inputStyle() {
  return {
    width:"100%",padding:"8px 10px",
    border:"1px solid #E0D8CE",borderRadius:8,
    fontSize:12,fontFamily:"'DM Sans',sans-serif",
    background:"#FAFAF8",color:"#1C1510",
  };
}
