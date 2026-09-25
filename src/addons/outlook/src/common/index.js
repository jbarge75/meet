const BASE_URL = window.__APP_CONFIG__?.BASE_URL || "https://meet.127.0.0.1.nip.io";
const APP_NAME = window.__APP_CONFIG__?.APP_NAME || "LaSuite Meet";
const ENABLE_SOURCE_TRACKING = window.__APP_CONFIG__?.ENABLE_SOURCE_TRACKING === "true";
const FEEDBACK_FORM = window.__APP_CONFIG__?.FEEDBACK_FORM || null;

// Comma-separated allowlist, e.g. "entreprise1.com,entreprise3.com".
// Empty/unset means the Polycom token is disabled for everyone (opt-in).
const POLYCOM_ALLOWED_DOMAINS = (window.__APP_CONFIG__?.POLYCOM_ALLOWED_DOMAINS || "")
  .split(",")
  .map((domain) => domain.trim().toLowerCase())
  .filter(Boolean);

function isPolycomDomainAllowed(email) {
  if (!email || POLYCOM_ALLOWED_DOMAINS.length === 0) return false;
  const domain = email.split("@")[1]?.toLowerCase();
  return POLYCOM_ALLOWED_DOMAINS.includes(domain);
}

module.exports = {
  BASE_URL,
  APP_NAME,
  ENABLE_SOURCE_TRACKING,
  FEEDBACK_FORM,
  POLYCOM_ALLOWED_DOMAINS,
  isPolycomDomainAllowed,
};
