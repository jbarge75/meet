const { APP_NAME } = require("./index");

function isOfficeReady() {
  return typeof Office !== "undefined" && Office?.context?.roamingSettings != null;
}

function applyAppName() {
  document.querySelectorAll("[data-app-name]").forEach((el) => {
    el.textContent = APP_NAME;
  });
}

/**
 * Resolves the actual body format of the item (HTML vs plain text), rather
 * than guessing it from the host platform. Desktop Outlook (Windows/Mac)
 * composes in HTML just as often as the web client does.
 */
function getIsHtmlBody(item) {
  return new Promise((resolve) => {
    item.body.getTypeAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        resolve(true);
        return;
      }
      resolve(result.value === Office.MailboxEnums.BodyType.Html);
    });
  });
}

module.exports = {
  isOfficeReady,
  applyAppName,
  getIsHtmlBody,
};
