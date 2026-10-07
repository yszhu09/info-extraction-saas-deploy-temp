import { NextRequest, NextResponse } from "next/server";

const noStorePaths = new Set(["/", "/compare", "/login"]);

function withNoStore(response: NextResponse, pathname: string) {
  if (noStorePaths.has(pathname)) {
    response.headers.set(
      "Cache-Control",
      "no-store, no-cache, max-age=0, must-revalidate",
    );
    response.headers.set("Pragma", "no-cache");
    response.headers.set("Expires", "0");
  }
  return response;
}

export function middleware(request: NextRequest) {
  if (request.nextUrl.pathname === "/login") {
    return withNoStore(NextResponse.next(), request.nextUrl.pathname);
  }
  if (request.cookies.get("info_extractor_auth")?.value === "ok") {
    return withNoStore(NextResponse.next(), request.nextUrl.pathname);
  }
  const loginUrl = request.nextUrl.clone();
  loginUrl.pathname = "/login";
  loginUrl.searchParams.set("next", request.nextUrl.pathname);
  return withNoStore(NextResponse.redirect(loginUrl), request.nextUrl.pathname);
}

export const config = {
  matcher: [
    "/((?!api|_next/static|_next/image|infowb-assets|favicon.ico|.*\\..*).*)",
  ],
};
