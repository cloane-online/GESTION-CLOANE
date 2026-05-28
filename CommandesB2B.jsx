import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import * as db from "./supabaseClient";

// ════════════════════════════════════════════════════════════════════════════
//   CONSTANTES — listes officielles (feuille 2 du tableau source)
// ════════════════════════════════════════════════════════════════════════════

const ETATS = [
  { val: "A commander",           label: "À commander",       color: "#A0620A", bg: "#FFF3DE" },
  { val: "Rupture Stock",         label: "Rupture stock",     color: "#9B2020", bg: "#FCEAEA" },
  { val: "DEPOT CLOANE",          label: "Dépôt CLOANE",      color: "#7A5A2A", bg: "#F5EBD8" },
  { val: "En Panier B2B",         label: "En panier B2B",     color: "#5C7AA0", bg: "#E5EFF7" },
  { val: "Commande validée",      label: "Commande validée",  color: "#2A7A3B", bg: "#E4F5E8" },
  { val: "Commande réceptionnée", label: "Cde réceptionnée",  color: "#5A2DA0", bg: "#F0E9FC" },
  { val: "Client prévenu",        label: "Client prévenu",    color: "#1B5E9B", bg: "#E3F0FC" },
  { val: "Produit encaissé",      label: "Produit encaissé",  color: "#1A6648", bg: "#DFF5EC" },
];

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
function buildClientMessage(row, magasin = "CLOANE SQUARE") {
  const prenom = (row.prenom || "").trim() || "";
  const marque = (row.marque || "").trim();
  return (
`Bonjour ${prenom},

Votre commande${marque ? " de la marque " + marque : ""} est bien arrivée en magasin.

Merci de nous recontacter afin de confirmer votre futur passage pour le récupérer.

En vous souhaitant une bonne journée,
L'équipe ${magasin}`
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
  articleCote:"", deCoteJusquau:"", dateRecup:""
};

// ── Helper components ──────────────────────────────────────────────────────

function badgeEtat(etat) {
  const s = ETATS.find(x => x.val === etat);
  if (!s) return etat ? <span style={{fontSize:10,color:"#A09080"}}>{etat}</span> : null;
  return (
    <span style={{
      display:"inline-block",padding:"3px 9px",borderRadius:20,
      fontSize:10,fontWeight:600,fontFamily:"'DM Sans',sans-serif",
      color:s.color,background:s.bg,whiteSpace:"nowrap",letterSpacing:"0.02em",
    }}>{s.label}</span>
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

export default function CommandesB2B({onBack}) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(false);
  const [brands, setBrands] = useState([...DEFAULT_BRANDS]);
  const [vendors, setVendors] = useState([...DEFAULT_VENDORS]);
  // URL B2B par marque/fournisseur (modifiable dans les paramètres)
  const [brandUrls, setBrandUrls] = useState({
    "CARHARTT": "https://b2b.carhartt-wip.com",
    "TOMMY HILFIGER": "https://b2b.tommy.com",
    "LACOSTE": "https://b2b.lacoste.com",
    "AIGLE": "https://pro.aigle.com",
    "REPLAY": "https://b2b.replay.it",
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

  // ── CHARGEMENT depuis Supabase + purge auto au montage ─────────────────────
  // Règle : on garde les 60 derniers jours sauf statut "commandé",
  //          et on supprime les "Produit encaissé" de plus de 15 jours.
  useEffect(() => {
    (async () => {
      try {
        const rows = await db.fetchOrders();
        const today = new Date();
        const sixtyAgo = addDays(today, -60);
        const fifteenAgo = addDays(today, -15);
        let removed = 0;
        const stale = [];
        const cleaned = rows.filter(r => {
          if (r.etat === "Produit encaissé") {
            const dv = parseDate(r.dateValid) || parseDate(r.date);
            if (dv && dv < fifteenAgo) { stale.push(r.id); removed++; return false; }
          }
          if (!STATUTS_COMMANDE.includes(r.etat)) {
            const dc = parseDate(r.date);
            if (dc && dc < sixtyAgo) { stale.push(r.id); removed++; return false; }
          }
          return true;
        });
        if (stale.length) db.deleteOrders(stale).catch(e => console.error("purge", e));
        setData(cleaned);
        setArchiveCount(removed);
      } catch (e) {
        console.error("Chargement Supabase échoué :", e);
        setLoadError(true);
      } finally {
        setLoading(false);
        setTimeout(() => setLoaded(true), 60);
      }
    })();
  }, []);

  // ── Filtering & sorting ──
  const filtered = useMemo(() => {
    let d = data;
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

  const stats = useMemo(() => ({
    total: data.length,
    aCommander: data.filter(r => r.etat === "A commander").length,
    clientPrevenu: data.filter(r => r.etat === "Client prévenu").length,
    rupture: data.filter(r => r.etat === "Rupture Stock").length,
    encaisse: data.filter(r => r.etat === "Produit encaissé").length,
  }), [data]);

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

  function deleteRow(id) {
    if (!confirm("Supprimer cette commande ?")) return;
    setData(d => d.filter(r => r.id !== id));
    setSelected(s => { const n = new Set(s); n.delete(id); return n; });
    db.deleteOrder(id).catch(e => console.error("deleteOrder", e));
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
    if (!form.nom && !form.modele && !form.marque) {
      alert("Veuillez renseigner au moins le nom du client, la marque ou le modèle.");
      return;
    }
    const today = fmtDate(new Date());
    if (editId) {
      const oldRow = data.find(r => r.id === editId);
      const statusChanged = oldRow && oldRow.etat !== form.etat;
      const newDateValid = statusChanged ? today : form.dateValid;
      updateRow(editId, {...form, dateValid: newDateValid});
      if (statusChanged && form.etat === "Client prévenu") {
        setTimeout(() => setTicketRow({
          ...form, id: editId, dateValid: newDateValid,
          articleCote: form.articleCote || [form.marque, form.modele, form.couleur && "Couleur : "+form.couleur, form.taille && "Taille : "+form.taille].filter(Boolean).join("\n"),
          deCoteJusquau: form.deCoteJusquau || fmtDate(addDays(new Date(), 7)),
        }), 200);
      }
      setShowEdit(false);
    } else {
      const draft = {...form, date: form.date || today};
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
      dateRecup: ticketRow.dateRecup || "",
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
  function setBrandUrl(name, url) {
    setBrandUrls(u => ({...u, [name]: url}));
  }
  // Ouvre le site B2B de la marque (bouton "Commander")
  function openOrderUrl(marque) {
    const url = (brandUrls[marque] || "").trim();
    if (!url) {
      alert(`Aucune URL B2B enregistrée pour « ${marque} ».\nAjoutez-la dans ⚙ Paramètres → Fournisseurs.`);
      setShowSettings(true); setSettingsTab("brands");
      return;
    }
    const full = /^https?:\/\//i.test(url) ? url : "https://" + url;
    window.open(full, "_blank", "noopener");
  }
  // Contact client : WhatsApp / Mail avec message prérempli
  function sendWhatsApp(row) {
    const intl = phoneIntl(row.tel);
    if (!intl) { alert("Numéro de téléphone manquant."); return; }
    const msg = encodeURIComponent(buildClientMessage(row));
    window.open(`https://wa.me/${intl}?text=${msg}`, "_blank", "noopener");
    setContactMenu(null);
  }
  function sendMail(row) {
    if (!row.email) { alert("Aucune adresse e-mail renseignée pour ce client (fiche commande)."); return; }
    const subject = encodeURIComponent("Votre commande CLOANE est arrivée");
    const body = encodeURIComponent(buildClientMessage(row));
    window.location.href = `mailto:${row.email}?subject=${subject}&body=${body}`;
    setContactMenu(null);
  }
  function sendSMS(row) {
    const tel = normalizePhone(row.tel);
    if (!tel) { alert("Numéro de téléphone manquant."); return; }
    const body = encodeURIComponent(buildClientMessage(row));
    window.location.href = `sms:${tel}?&body=${body}`;
    setContactMenu(null);
  }
  function addVendor() {
    const name = newVendorInput.trim().toUpperCase();
    if (!name || vendors.includes(name)) return;
    setVendors(v => [...v, name].sort());
    setNewVendorInput("");
  }
  function removeVendor(name) {
    if (!confirm(`Supprimer le conseiller "${name}" de la liste ?`)) return;
    setVendors(v => v.filter(x => x !== name));
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
      "Jour récupération": fmtDateFr(r.dateRecup),
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
          padding:24px 20px; line-height:1.55;
          border:1px dashed #D8CCBE; border-radius:6px;
          width:280px; margin:0 auto;
        }
        @media print {
          body * { visibility:hidden !important; }
          .ticket-printable, .ticket-printable * { visibility:visible !important; }
          .ticket-printable {
            position:fixed !important; top:0 !important; left:0 !important;
            width:76mm !important; padding:8mm !important;
            border:none !important; background:#fff !important;
          }
          @page { size:76mm auto; margin:0; }
        }
      `}</style>

      {/* ── HEADER ─────────────────────────────────────────────────────── */}
      <div style={{
        position:"sticky",top:0,zIndex:50,
        background:"rgba(244,240,232,0.96)",backdropFilter:"blur(12px)",
        borderBottom:"1px solid rgba(200,169,110,0.2)",
        padding:"0 24px",display:"flex",alignItems:"center",
        justifyContent:"space-between",height:60,gap:12,flexWrap:"wrap",
      }}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          {onBack ? (
            <>
              <button onClick={onBack} style={{
                all:"unset",cursor:"pointer",display:"flex",alignItems:"center",gap:6,
                color:"#8A7A6A",fontSize:11,fontFamily:"'DM Sans',sans-serif",
                letterSpacing:"0.12em",textTransform:"uppercase",
              }}>
                <span style={{fontSize:16}}>←</span>Menu
              </button>
              <div style={{width:1,height:18,background:"#D8CCBE"}}/>
            </>
          ) : (
            <div style={{
              width:30,height:30,background:"#1C1510",borderRadius:"50%",
              display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,
            }}>
              <span style={{color:"#C8A96E",fontFamily:"'Cormorant Garamond',serif",fontSize:13,fontWeight:500}}>C</span>
            </div>
          )}
          <div style={{
            fontFamily:"'Cormorant Garamond',serif",fontSize:18,fontWeight:400,
            color:"#1C1510",letterSpacing:"0.08em",textTransform:"uppercase",
          }}>
            Commandes Clients B2B
          </div>
          <span style={{
            fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"#A09080",
            background:"#EDE4D5",padding:"2px 9px",borderRadius:20,
          }}>{filtered.length} / {data.length}</span>
        </div>

        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          {backupAvailable && (
            <button onClick={exportBackup} className="pulse" style={{
              background:"#FFF8EB",border:"1px solid #C8A96E",borderRadius:8,
              padding:"6px 14px",fontSize:11,cursor:"pointer",
              fontFamily:"'DM Sans',sans-serif",color:"#7A5A2A",
              display:"flex",alignItems:"center",gap:6,fontWeight:600,
            }} title={`Sauvegarde Excel quotidienne — ${todayDayName}`}>
              💾 Sauvegarde du jour
            </button>
          )}
          <button onClick={()=>setShowSettings(true)} style={{
            background:"#F5F0E8",border:"1px solid #DDD4C8",borderRadius:8,
            padding:"6px 10px",fontSize:13,cursor:"pointer",color:"#5A4030",
          }} title="Gérer les marques et conseillers">⚙</button>
          <button onClick={openNew} style={{
            background:"#1C1510",color:"#E8DED0",border:"none",borderRadius:8,
            padding:"7px 16px",fontSize:12,cursor:"pointer",
            fontFamily:"'DM Sans',sans-serif",letterSpacing:"0.06em",
            display:"flex",alignItems:"center",gap:6,
          }}><span style={{fontSize:14}}>+</span> Nouvelle commande</button>
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
            <span><b>{archiveCount}</b> commande(s) archivée(s) automatiquement (purge des +60 jours hors statut commandé, et des produits encaissés +15 jours).</span>
          </div>
        )}

        {/* ── STATS ROW ──────────────────────────────────────────────── */}
        <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap"}}>
          {[
            {label:"Total",val:stats.total,color:"#8A7A6A",bg:"#F5F0E8",etat:null},
            {label:"À commander",val:stats.aCommander,color:"#A0620A",bg:"#FFF3DE",etat:"A commander"},
            {label:"Client prévenu",val:stats.clientPrevenu,color:"#1B5E9B",bg:"#E3F0FC",etat:"Client prévenu"},
            {label:"Rupture stock",val:stats.rupture,color:"#9B2020",bg:"#FCEAEA",etat:"Rupture Stock"},
            {label:"Encaissé",val:stats.encaisse,color:"#1A6648",bg:"#DFF5EC",etat:"Produit encaissé"},
          ].map(s => (
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
          <DropDown options={vendors} value={filters.vendeur}
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
                      <td style={TD_STYLE}>{badgeEtat(row.etat)||<span style={{color:"#C0B0A0",fontSize:10}}>—</span>}</td>
                      <td style={{...TD_STYLE,color:"#8A6A5A",fontSize:11,fontWeight:500}}>{row.vendeur}</td>
                      <td style={{...TD_STYLE,whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                        {row.marque && (
                          <button onClick={()=>openOrderUrl(row.marque)}
                            title={brandUrls[row.marque] ? `Commander sur ${brandUrls[row.marque]}` : "URL B2B à renseigner dans les paramètres"}
                            style={{
                              fontSize:11,padding:"3px 9px",
                              background: brandUrls[row.marque] ? "#1C1510" : "#F0EBE3",
                              border: brandUrls[row.marque] ? "1px solid #1C1510" : "1px solid #DDD4C8",
                              borderRadius:6,cursor:"pointer",
                              color: brandUrls[row.marque] ? "#E8DED0" : "#9A8A7A",
                              marginRight:4,fontFamily:"'DM Sans',sans-serif",
                            }}>🛒 Commander</button>
                        )}
                        {row.etat === "Client prévenu" && (
                          <button onClick={()=>setTicketRow({
                            ...row,
                            articleCote: row.articleCote || [row.marque,row.modele,row.couleur&&"Couleur : "+row.couleur,row.taille&&"Taille : "+row.taille].filter(Boolean).join("\n"),
                            deCoteJusquau: row.deCoteJusquau || fmtDate(addDays(new Date(),7)),
                          })} title="Voir / imprimer le ticket caisse" style={{
                            fontSize:11,padding:"3px 8px",background:"#E3F0FC",border:"1px solid #B0CDE6",
                            borderRadius:6,cursor:"pointer",color:"#1B5E9B",marginRight:4,
                          }}>🎫</button>
                        )}
                        <button onClick={()=>openEdit(row)} title="Modifier" style={{
                          fontSize:11,padding:"3px 9px",background:"#F5F0E8",border:"1px solid #DDD4C8",
                          borderRadius:6,cursor:"pointer",color:"#5A4030",marginRight:4,
                        }}>✏</button>
                        <button onClick={()=>deleteRow(row.id)} title="Supprimer" style={{
                          fontSize:11,padding:"3px 9px",background:"#FDF0F0",border:"1px solid #F0D0D0",
                          borderRadius:6,cursor:"pointer",color:"#9B2020",
                        }}>✕</button>
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
              <FormField label="Nom client">
                <input value={form.nom||""} onChange={e=>setForm(p=>({...p,nom:e.target.value}))} style={inputStyle()}/>
              </FormField>
              <FormField label="Prénom">
                <input value={form.prenom||""} onChange={e=>setForm(p=>({...p,prenom:e.target.value}))} style={inputStyle()}/>
              </FormField>
              <FormField label="Téléphone">
                <input value={form.tel||""} onChange={e=>setForm(p=>({...p,tel:e.target.value}))} style={inputStyle()} placeholder="06 12 34 56 78"/>
              </FormField>
              <FormField label="E-mail">
                <input type="email" value={form.email||""} onChange={e=>setForm(p=>({...p,email:e.target.value}))} style={inputStyle()} placeholder="client@email.com"/>
              </FormField>
              <FormField label="Conseiller">
                <select value={form.vendeur||""} onChange={e=>setForm(p=>({...p,vendeur:e.target.value}))} style={inputStyle()}>
                  <option value="">—</option>
                  {vendors.map(v => <option key={v} value={v}>{v}</option>)}
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
                <input value={form.modele||""} onChange={e=>setForm(p=>({...p,modele:e.target.value}))} style={inputStyle()}/>
              </FormField>
              <FormField label="Réf. interne">
                <input value={form.refInt||""} onChange={e=>setForm(p=>({...p,refInt:e.target.value}))} style={inputStyle()} placeholder="Saisie libre"/>
              </FormField>
              <FormField label="Réf. fournisseur">
                <input value={form.refFourn||""} onChange={e=>setForm(p=>({...p,refFourn:e.target.value}))} style={inputStyle()} placeholder="Saisie libre"/>
              </FormField>
              <FormField label="Couleur">
                {(() => {
                  const isOther = form.couleur && !COLORS_TOP15.includes(form.couleur);
                  return (
                    <>
                      <select
                        value={isOther ? "__autre__" : (form.couleur || "")}
                        onChange={e => {
                          const v = e.target.value;
                          if (v === "__autre__") setForm(p => ({...p, couleur: " "})); // marque "autre", champ libre
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
                          onChange={e => setForm(p => ({...p, couleur: e.target.value || " "}))}
                          placeholder="Préciser la couleur"
                          style={{...inputStyle(), marginTop:6}}/>
                      )}
                    </>
                  );
                })()}
              </FormField>
              <FormField label="Taille">
                <input value={form.taille||""} onChange={e=>setForm(p=>({...p,taille:e.target.value}))} style={inputStyle()}/>
              </FormField>
            </div>
            <FormField label="Commentaire">
              <textarea value={form.commentaire||""} onChange={e=>setForm(p=>({...p,commentaire:e.target.value}))} rows={3} style={{...inputStyle(),resize:"vertical",minHeight:60}}/>
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
                  {vendors.map(v => <option key={v} value={v}>{v}</option>)}
                </select>
              </FormField>
            </div>

            {/* Jour de récupération souhaité — calendrier + raccourcis */}
            <div style={{
              background:"#F0F5FB",border:"1px solid #D6E4F0",borderRadius:10,
              padding:"12px 14px",marginBottom:18,
            }}>
              <label style={{
                display:"block",fontSize:10,color:"#1B5E9B",
                fontFamily:"'DM Sans',sans-serif",fontWeight:700,
                textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:8,
              }}>📅 Jour de récupération de la commande</label>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                <input type="date" value={ticketRow.dateRecup||""}
                  onChange={e=>setTicketRow(r=>({...r,dateRecup:e.target.value}))}
                  style={{...inputStyle(),maxWidth:180}}/>
                {[
                  {l:"Par défaut (J+7)", d:7},
                  {l:"+3 j", d:3},
                  {l:"+14 j", d:14},
                ].map(p => (
                  <button key={p.l} onClick={()=>setTicketRow(r=>({...r,dateRecup:fmtDate(addDays(new Date(),p.d))}))}
                    style={{
                      fontSize:11,padding:"6px 10px",background:"#FFF",
                      border:"1px solid #B0CDE6",borderRadius:7,cursor:"pointer",
                      color:"#1B5E9B",fontFamily:"'DM Sans',sans-serif",
                    }}>{p.l}</button>
                ))}
              </div>
            </div>
            <FormField label="Quel article reste de côté ?">
              <textarea value={ticketRow.articleCote||""}
                onChange={e=>setTicketRow(r=>({...r,articleCote:e.target.value}))}
                rows={4} style={{...inputStyle(),resize:"vertical",minHeight:80,fontFamily:"'Courier New',monospace"}}/>
            </FormField>

            <div style={{
              marginTop:18,marginBottom:18,
              background:"#F8F3EB",padding:"14px",borderRadius:10,
              fontSize:10,color:"#8A7A6A",fontFamily:"'DM Sans',sans-serif",
              textAlign:"center",letterSpacing:"0.06em",textTransform:"uppercase",
            }}>↓ Aperçu du ticket caisse ↓</div>

            <div className="ticket-printable">
              <div className="ticket-paper">
                <div style={{textAlign:"center",fontSize:11,letterSpacing:"0.1em",borderBottom:"1px dashed #8A7A6A",paddingBottom:6,marginBottom:10}}>
                  ━━━━━━━━━━━━━━━━━━<br/>
                  <b style={{fontSize:13}}>COMMANDE CLIENT</b><br/>
                  ━━━━━━━━━━━━━━━━━━
                </div>
                <div style={{fontSize:13,lineHeight:1.7}}>
                  <b style={{fontSize:14}}>{(ticketRow.nom||"").toUpperCase()} {ticketRow.prenom||""}</b><br/>
                  {ticketRow.tel ? fmtPhone(ticketRow.tel) : "—"}
                </div>
                <div style={{margin:"14px 0",borderTop:"1px dashed #C8B89A",paddingTop:12,fontSize:12,textAlign:"center"}}>
                  <div style={{fontSize:10,color:"#6A5A4A",marginBottom:4,letterSpacing:"0.08em"}}>DE CÔTÉ JUSQU'AU</div>
                  <div style={{fontSize:16,fontWeight:700,letterSpacing:"0.06em"}}>
                    {ticketRow.deCoteJusquau ? fmtDateFr(ticketRow.deCoteJusquau) : "__ / __ / ____"}
                  </div>
                </div>
                {ticketRow.dateRecup && (
                  <div style={{margin:"10px 0",fontSize:12,textAlign:"center"}}>
                    <div style={{fontSize:10,color:"#6A5A4A",marginBottom:3,letterSpacing:"0.08em"}}>JOUR DE RÉCUPÉRATION</div>
                    <b style={{fontSize:14}}>{fmtDateFr(ticketRow.dateRecup)}</b>
                  </div>
                )}
                <div style={{margin:"10px 0",fontSize:11,textAlign:"center"}}>
                  <div style={{color:"#6A5A4A",fontSize:10,letterSpacing:"0.08em",marginBottom:3}}>CONSEILLER</div>
                  <b style={{fontSize:12}}>{ticketRow.vendeur||"—"}</b>
                </div>
                {ticketRow.articleCote && (
                  <div style={{marginTop:14,paddingTop:10,borderTop:"1px dashed #C8B89A",fontSize:11,whiteSpace:"pre-line",textAlign:"center"}}>
                    <div style={{fontSize:9,color:"#6A5A4A",letterSpacing:"0.08em",marginBottom:4}}>ARTICLE RÉSERVÉ</div>
                    {ticketRow.articleCote}
                  </div>
                )}
                <div style={{textAlign:"center",marginTop:14,fontSize:9,color:"#8A7A6A",letterSpacing:"0.16em"}}>
                  ━━━━━━━━━━━━━━━━━━<br/>
                  C L O A N E
                </div>
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
                {key:"vendors",label:`Conseillers (${vendors.length})`},
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
                    Renseignez l'URL B2B de chaque fournisseur : elle sera utilisée par le bouton 🛒 Commander du tableau.
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:6}}>
                    {brands.map(b => (
                      <div key={b} style={{
                        display:"flex",alignItems:"center",gap:8,
                        background:"#FAFAF8",border:"1px solid #EDE4D5",borderRadius:10,
                        padding:"7px 8px 7px 12px",
                      }}>
                        <span style={{
                          fontSize:11,fontFamily:"'DM Sans',sans-serif",fontWeight:600,
                          color:"#4A2C1A",minWidth:120,flexShrink:0,
                        }}>{b}</span>
                        <input
                          value={brandUrls[b] || ""}
                          onChange={e=>setBrandUrl(b, e.target.value)}
                          placeholder="URL B2B (ex : b2b.marque.com)"
                          style={{
                            flex:1,padding:"6px 9px",fontSize:11,
                            fontFamily:"'DM Sans',sans-serif",
                            border:"1px solid #E0D8CE",borderRadius:7,
                            background:"#FFF",color:"#1C1510",minWidth:80,
                          }}/>
                        {brandUrls[b] && (
                          <button onClick={()=>openOrderUrl(b)} title="Tester le lien" style={{
                            background:"#E3F0FC",border:"1px solid #B0CDE6",borderRadius:6,
                            padding:"5px 8px",fontSize:11,cursor:"pointer",color:"#1B5E9B",flexShrink:0,
                          }}>↗</button>
                        )}
                        <button onClick={()=>removeBrand(b)} title="Supprimer" style={{
                          background:"#FDEEEE",border:"none",borderRadius:"50%",
                          width:22,height:22,cursor:"pointer",color:"#9B2020",fontSize:13,lineHeight:1,flexShrink:0,
                        }}>×</button>
                      </div>
                    ))}
                  </div>
                </>
              ) : (
                <>
                  <div style={{display:"flex",gap:8,marginBottom:12}}>
                    <input value={newVendorInput} onChange={e=>setNewVendorInput(e.target.value)}
                      onKeyDown={e=>e.key==="Enter"&&addVendor()}
                      placeholder="Nouveau conseiller (sera mis en majuscules)"
                      style={inputStyle()}/>
                    <button onClick={addVendor} style={{
                      background:"#1C1510",color:"#E8DED0",border:"none",borderRadius:8,
                      padding:"0 16px",fontSize:12,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",
                    }}>+ Ajouter</button>
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                    {vendors.map(v => (
                      <span key={v} style={{
                        display:"inline-flex",alignItems:"center",gap:6,
                        background:"#F5F0E8",color:"#1C1510",fontSize:11,
                        fontFamily:"'DM Sans',sans-serif",
                        padding:"4px 4px 4px 12px",borderRadius:20,
                        border:"1px solid #E0D8CE",
                      }}>
                        {v}
                        <button onClick={()=>removeVendor(v)} style={{
                          background:"#FDEEEE",border:"none",borderRadius:"50%",
                          width:18,height:18,cursor:"pointer",color:"#9B2020",fontSize:11,lineHeight:1,
                        }}>×</button>
                      </span>
                    ))}
                  </div>
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

// ── Inline helper components ────────────────────────────────────────────────

function Modal({children, onClose, title, maxWidth=640}) {
  return (
    <div onClick={onClose} style={{
      position:"fixed",inset:0,zIndex:1000,
      background:"rgba(28,21,16,0.55)",backdropFilter:"blur(4px)",
      display:"flex",alignItems:"center",justifyContent:"center",padding:16,
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
