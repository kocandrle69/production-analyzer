import { NextResponse } from "next/server";

export function middleware(request) {
  const auth = request.headers.get("authorization") ?? "";

  if (auth.startsWith("Basic ")) {
    const [user, pass] = atob(auth.slice(6)).split(":");
    const validUser = process.env.AUTH_USER || "admin";
    const validPass = process.env.AUTH_PASSWORD || "";

    if (validPass && user === validUser && pass === validPass) {
      return NextResponse.next();
    }
  }

  return new NextResponse("Unauthorized", {
    status: 401,
    headers: { "WWW-Authenticate": 'Basic realm="Production Analyzer"' },
  });
}

export const config = {
  matcher: ["/((?!_vercel).*)"],
};
