import { NextResponse } from "next/server";

export async function POST(request: Request) {
  const body = (await request.json().catch(() => ({}))) as {
    password?: string;
  };
  const expectedPassword = process.env.APP_PASSWORD || "demo123";
  if (body.password !== expectedPassword) {
    return NextResponse.json(
      { ok: false, error: "密码不正确" },
      { status: 401 },
    );
  }
  const response = NextResponse.json({ ok: true });
  response.cookies.set("info_extractor_auth", "ok", {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/",
    maxAge: 60 * 60 * 24 * 7,
  });
  return response;
}
