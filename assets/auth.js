// 簡易ID/PWゲート（SHA-256ハッシュ照合）
// 平文は埋め込まない。ハッシュとソルトのみ。

const AUTH_HASH = "a3858dfadb73cb0a3244716122664cf3924c9b91549b3e3018d8a67e67b7c7f9";
const AUTH_SALT = "tamaki-zaisan-2026";
const SESSION_KEY = "tamaki-zaisan-auth";

async function sha256Hex(str) {
  const buf = new TextEncoder().encode(str);
  const digest = await crypto.subtle.digest("SHA-256", buf);
  return Array.from(new Uint8Array(digest))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

async function tryLogin(id, pw) {
  const combined = `${AUTH_SALT}:${id.trim()}:${pw.trim()}`;
  const h = await sha256Hex(combined);
  return h === AUTH_HASH;
}

function isAuthed() {
  return sessionStorage.getItem(SESSION_KEY) === "1";
}

function setAuthed() {
  sessionStorage.setItem(SESSION_KEY, "1");
}

function clearAuth() {
  sessionStorage.removeItem(SESSION_KEY);
}
