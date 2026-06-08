import { useEffect, useState } from "react";
import CommandesB2B from "./CommandesB2B.jsx";
import LoginScreen from "./LoginScreen.jsx";
import * as db from "./supabaseClient";

export default function App() {
  const [session, setSession] = useState(null);
  const [profile, setProfile] = useState(null);
  const [loading, setLoading] = useState(true);
  const [profileError, setProfileError] = useState("");

  useEffect(() => {
    // Charger session existante au démarrage
    db.getSession().then(s => {
      setSession(s);
      if (s) loadProfile(s.user.id);
      else setLoading(false);
    });
    // Écouter les changements (login/logout)
    const { data: { subscription } } = db.onAuthChange((s) => {
      setSession(s);
      if (s) loadProfile(s.user.id);
      else { setProfile(null); setLoading(false); }
    });
    return () => subscription?.unsubscribe?.();
  }, []);

  async function loadProfile(userId) {
    setLoading(true); setProfileError("");
    try {
      const p = await db.fetchProfile(userId);
      setProfile(p);
    } catch (e) {
      console.error(e);
      setProfileError("Profil utilisateur introuvable. Contactez un administrateur.");
    } finally {
      setLoading(false);
    }
  }

  if (loading) {
    return (
      <div style={{
        minHeight:"100vh",background:"#F4F0E8",
        display:"flex",alignItems:"center",justifyContent:"center",
        fontFamily:"'DM Sans',sans-serif",color:"#8A7A6A",fontSize:13,
      }}>⏳ Connexion…</div>
    );
  }

  if (!session) {
    return <LoginScreen onLoggedIn={() => {}}/>;
  }

  if (profileError) {
    return (
      <div style={{
        minHeight:"100vh",background:"#F4F0E8",
        display:"flex",alignItems:"center",justifyContent:"center",padding:20,
        fontFamily:"'DM Sans',sans-serif",
      }}>
        <div style={{
          background:"#FFF",border:"1px solid #F0C8C8",borderRadius:14,
          padding:"24px 28px",maxWidth:400,textAlign:"center",
        }}>
          <div style={{color:"#9B2020",fontSize:14,marginBottom:14}}>{profileError}</div>
          <button onClick={()=>db.signOut()} style={{
            padding:"9px 16px",background:"#1C1510",color:"#E8DED0",border:"none",
            borderRadius:8,fontSize:12,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",
          }}>Déconnexion</button>
        </div>
      </div>
    );
  }

  return <CommandesB2B session={session} profile={profile} onSignOut={() => db.signOut()}/>;
}
