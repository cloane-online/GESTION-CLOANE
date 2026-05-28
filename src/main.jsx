import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { supabase } from "./supabaseClient";

function Login() {
  const [email, setEmail] = React.useState("");
  const [password, setPassword] = React.useState("");
  const [error, setError] = React.useState("");

  const handleLogin = async (e) => {
    e.preventDefault();

    const { error } = await supabase.auth.signInWithPassword({
      email,
      password,
    });

    if (error) {
      setError(error.message);
    } else {
      window.location.reload();
    }
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        background: "#f5f5f5",
        fontFamily: "Arial",
      }}
    >
      <form
        onSubmit={handleLogin}
        style={{
          background: "white",
          padding: 30,
          borderRadius: 12,
          width: 320,
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
        }}
      >
        <h2 style={{ marginBottom: 20 }}>
          Connexion CLOANE
        </h2>

        <input
          type="email"
          placeholder="E-mail"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          style={{
            width: "100%",
            padding: 10,
            marginBottom: 12,
          }}
        />

        <input
          type="password"
          placeholder="Mot de passe"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          style={{
            width: "100%",
            padding: 10,
            marginBottom: 12,
          }}
        />

        <button
          type="submit"
          style={{
            width: "100%",
            padding: 10,
            cursor: "pointer",
          }}
        >
          Se connecter
        </button>

        {error && (
          <p style={{ color: "red", marginTop: 12 }}>
            {error}
          </p>
        )}
      </form>
    </div>
  );
}

async function start() {
  const {
    data: { session },
  } = await supabase.auth.getSession();

  const root = ReactDOM.createRoot(
    document.getElementById("root")
  );

  if (session) {
    root.render(
      <React.StrictMode>
        <App />
      </React.StrictMode>
    );
  } else {
    root.render(
      <React.StrictMode>
        <Login />
      </React.StrictMode>
    );
  }
}

start();
