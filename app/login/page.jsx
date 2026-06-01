"use client";

import { useEffect, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import { login } from "@/lib/api";

export default function LoginPage() {
  const router = useRouter();
  const emailRef = useRef(null);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    emailRef.current?.focus();
  }, []);

  async function handleSubmit(e) {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const { token } = await login(email.trim(), password);
      localStorage.setItem("wbr_token", token);
      router.replace("/");
    } catch (err) {
      setError(err.message || "Login failed");
      setLoading(false);
    }
  }

  return (
    <div className="login-screen">
      <img src="/kutlerri-logo.png" alt="Kutlerri" className="login-logo" />

      <form className="login-card" onSubmit={handleSubmit}>
        <div className="login-title">RASA Analytics</div>
        <div className="login-sub">Weekly Business Review Dashboard</div>

        <label className="login-label" htmlFor="login-email">
          Email
        </label>
        <input
          id="login-email"
          ref={emailRef}
          type="email"
          className="login-input"
          placeholder=""
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          required
          autoComplete="email"
        />

        <label className="login-label" htmlFor="login-password">
          Password
        </label>
        <input
          id="login-password"
          type="password"
          className="login-input"
          placeholder=""
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          required
          autoComplete="current-password"
        />

        {error && <div className="login-error">{error}</div>}

        <button type="submit" className="login-btn" disabled={loading}>
          {loading ? "Signing in…" : "Sign In"}
        </button>
      </form>
    </div>
  );
}
