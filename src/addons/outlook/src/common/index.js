const BASE_URL = window.__APP_CONFIG__?.BASE_URL || "https://meet.127.0.0.1.nip.io";
const APP_NAME = window.__APP_CONFIG__?.APP_NAME || "LaSuite Meet";
const ENABLE_SOURCE_TRACKING = window.__APP_CONFIG__?.ENABLE_SOURCE_TRACKING === "true";
const FEEDBACK_FORM = window.__APP_CONFIG__?.FEEDBACK_FORM || null;

// Set by whoever deploys the manifest (e.g. ?polycom=true appended to the
// taskpane/commands source URLs), so an org can opt into the hidden Polycom
// token without depending on the central server's own config.js.
const POLYCOM_ENABLED = new URLSearchParams(window.location.search).get("polycom") === "true";

module.exports = {
  BASE_URL,
  APP_NAME,
  ENABLE_SOURCE_TRACKING,
  FEEDBACK_FORM,
  POLYCOM_ENABLED,
};
