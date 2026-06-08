import { useState } from "react";
import * as db from "./supabaseClient";

export default function LoginScreen({ onLoggedIn }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [forgotMode, setForgotMode] = useState(false);
  const [resetSent, setResetSent] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setError("");
    if (forgotMode) {
      if (!email) { setError("Saisissez votre adresse e-mail."); return; }
      setLoading(true);
      try {
        await db.resetPassword(email);
        setResetSent(true);
      } catch (err) {
        setError("Impossible d'envoyer le lien : " + (err.message || err));
      } finally { setLoading(false); }
      return;
    }
    if (!email || !password) { setError("E-mail et mot de passe requis."); return; }
    setLoading(true);
    try {
      await db.signIn(email.trim(), password);
      onLoggedIn?.();
    } catch (err) {
      setError("Identifiants incorrects. Réessayez ou utilisez « Mot de passe oublié ».");
    } finally { setLoading(false); }
  }

  return (
    <div style={{
      minHeight:"100vh",background:"linear-gradient(135deg, #F4F0E8 0%, #EDE4D5 100%)",
      display:"flex",alignItems:"center",justifyContent:"center",padding:20,
      fontFamily:"'DM Sans',sans-serif",
    }}>
      <div style={{
        background:"#FFFCF8",borderRadius:18,padding:"40px 36px 32px",
        width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(28,21,16,0.12)",
        border:"1px solid #EDE4D5",
      }}>
        <div style={{textAlign:"center",marginBottom:30}}>
          <div style={{
            fontFamily:"'Cormorant Garamond',serif",fontSize:38,fontWeight:300,
            letterSpacing:"0.22em",color:"#1C1510",lineHeight:1,
          }}>CLOANE</div>
          <div style={{
            fontSize:10,letterSpacing:"0.3em",color:"#8A7A6A",marginTop:8,
            textTransform:"uppercase",
          }}>Gestion des commandes</div>
        </div>

        {resetSent ? (
          <div style={{textAlign:"center",padding:"20px 0"}}>
            <div style={{fontSize:36,marginBottom:12}}>✉️</div>
            <div style={{fontSize:14,color:"#1C1510",marginBottom:8,fontWeight:600}}>
              Lien envoyé !
            </div>
            <div style={{fontSize:12,color:"#6A5A4A",lineHeight:1.5,marginBottom:20}}>
              Un e-mail vient de partir vers <b>{email}</b>.<br/>
              Cliquez sur le lien pour définir un nouveau mot de passe.
            </div>
            <button onClick={()=>{setForgotMode(false);setResetSent(false);setEmail("");setPassword("");}}
              style={btnSecondary()}>← Retour à la connexion</button>
          </div>
        ) : (
          <form onSubmit={submit}>
            <Label>E-mail</Label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)}
              autoComplete="email" placeholder="prenom@email.com"
              style={inputStyle()} disabled={loading}/>

            {!forgotMode && (<>
              <Label>Mot de passe</Label>
              <input type="password" value={password} onChange={e=>setPassword(e.target.value)}
                autoComplete="current-password" style={inputStyle()} disabled={loading}/>
            </>)}

            {error && (
              <div style={{
                background:"#FCEAEA",border:"1px solid #F0C8C8",borderRadius:8,
                padding:"10px 14px",fontSize:12,color:"#9B2020",marginBottom:14,
              }}>{error}</div>
            )}

            <button type="submit" disabled={loading} style={btnPrimary(loading)}>
              {loading ? "…" : (forgotMode ? "Recevoir un lien de réinitialisation" : "Se connecter")}
            </button>

            <div style={{textAlign:"center",marginTop:18}}>
              <button type="button" onClick={()=>{setForgotMode(m=>!m);setError("");}}
                style={{
                  all:"unset",cursor:"pointer",fontSize:12,color:"#8A7A6A",
                  textDecoration:"underline",textDecorationStyle:"dotted",
                }}>
                {forgotMode ? "← Retour à la connexion" : "Mot de passe oublié ?"}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}

function Label({children}) {
  return (
    <div style={{
      fontSize:10,letterSpacing:"0.12em",color:"#8A7A6A",
      textTransform:"uppercase",fontWeight:600,marginBottom:6,marginTop:14,
    }}>{children}</div>
  );
}
function inputStyle() {
  return {
    width:"100%",padding:"11px 14px",fontSize:14,
    fontFamily:"'DM Sans',sans-serif",color:"#1C1510",
    border:"1px solid #DDD4C8",borderRadius:9,background:"#FFF",
    outline:"none",boxSizing:"border-box",
  };
}
function btnPrimary(loading) {
  return {
    width:"100%",padding:"13px 16px",fontSize:13,fontWeight:600,
    fontFamily:"'DM Sans',sans-serif",letterSpacing:"0.06em",
    background: loading ? "#888" : "#1C1510",color:"#E8DED0",
    border:"none",borderRadius:9,cursor: loading ? "default" : "pointer",
    marginTop:18,
  };
}
function btnSecondary() {
  return {
    padding:"10px 16px",fontSize:12,fontFamily:"'DM Sans',sans-serif",
    background:"#F5F0E8",color:"#5A4030",border:"1px solid #DDD4C8",
    borderRadius:8,cursor:"pointer",
  };
}
