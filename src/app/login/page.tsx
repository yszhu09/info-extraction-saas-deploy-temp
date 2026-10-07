"use client";

import { useState } from "react";

export default function LoginPage() {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function login(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    setError("");
    const response = await fetch("/api/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password }),
    });
    setBusy(false);
    if (!response.ok) {
      setError("密码不正确");
      return;
    }
    const params = new URLSearchParams(window.location.search);
    window.location.href = params.get("next") || "/";
  }

  return (
    <main className="grid min-h-screen place-items-center bg-slate-950 px-4 text-white">
      <form
        onSubmit={login}
        className="w-full max-w-md rounded-3xl border border-white/10 bg-white p-8 text-slate-950 shadow-2xl"
      >
        <p className="text-sm font-semibold uppercase tracking-[0.25em] text-cyan-600">
          Private MVP
        </p>
        <h1 className="mt-2 text-3xl font-bold">信息提取器</h1>
        <p className="mt-3 text-sm leading-6 text-slate-500">
          请输入访问密码。该页面仅限授权用户访问。
        </p>
        <input
          className="mt-6 w-full rounded-2xl border border-slate-200 px-4 py-3 outline-none focus:border-cyan-500"
          type="password"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          placeholder="访问密码"
        />
        {error ? <p className="mt-3 text-sm text-red-600">{error}</p> : null}
        <button
          className="mt-5 w-full rounded-full bg-cyan-600 px-5 py-3 font-semibold text-white disabled:opacity-50"
          disabled={busy}
        >
          {busy ? "登录中…" : "进入工具"}
        </button>
      </form>
    </main>
  );
}
