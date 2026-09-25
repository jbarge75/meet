/* global Office */
const { APP_NAME, FEEDBACK_FORM, POLYCOM_ENABLED } = require("../common");
const { applyAppName, getIsHtmlBody } = require("../common/helpers");
const { buildMeetingMessage } = require("../common/messageBuilder");
const { initI18n, t, translateUI } = require("../common/i18n");
const { isMeetingAlreadyAdded, removeMeetingLink } = require("../common/meetingDetector");
// TEST ONLY — bypasses ProConnect auth and real room creation. Not for merge.
const { TEST_ROOM_DATA } = require("../common/testFixture");

// ── Views ────────────────────────────────────────────────────

function showView(name) {
  document.getElementById("view-loading").style.display = "none";
  document.getElementById("view-unauth").style.display = "none";
  document.getElementById("view-auth").style.display = "none";
  document.getElementById(`view-${name}`).style.display = "block";

  if (name === "auth") {
    _refreshMeetingButtonState();
  }
}

// ── Button state ─────────────────────────────────────────────

function _showAddButton() {
  document.getElementById("btn-generate").style.display = "block";
  document.getElementById("btn-remove").style.display = "none";
}

function _showRemoveButton() {
  document.getElementById("btn-generate").style.display = "none";
  document.getElementById("btn-remove").style.display = "block";
}

function _setButtonLoading() {
  const btn = document.getElementById("btn-generate");
  btn.disabled = true;
  btn.textContent = t("meeting.generating");
}

function _setButtonIdle() {
  const btn = document.getElementById("btn-generate");
  btn.disabled = false;
  btn.textContent = t("meeting.add_meeting", { app_name: APP_NAME });
}

function _setRemoveLoading() {
  const btn = document.getElementById("btn-remove");
  btn.disabled = true;
  btn.textContent = t("meeting.removing");
}

function _setRemoveIdle() {
  const btn = document.getElementById("btn-remove");
  btn.disabled = false;
  btn.textContent = t("meeting.remove_meeting", { app_name: APP_NAME });
}

function _renderPolycomStatus() {
  const badge = document.getElementById("polycom-status-badge");
  badge.textContent = t(POLYCOM_ENABLED ? "polycom.enabled" : "polycom.disabled");
  badge.classList.toggle("is-off", !POLYCOM_ENABLED);
}

function _refreshMeetingButtonState() {
  const item = Office.context.mailbox.item;
  if (!item) return;
  isMeetingAlreadyAdded(item).then((alreadyAdded) => {
    if (alreadyAdded) {
      _showRemoveButton();
    } else {
      _showAddButton();
    }
  });
}

// ── Meeting ──────────────────────────────────────────────────

function generateMeetingLink() {
  // TEST ONLY — ProConnect auth skipped, uses fixture room data directly.
  _setButtonLoading();

  const item = Office.context.mailbox.item;

  Promise.all([Promise.resolve(TEST_ROOM_DATA), getIsHtmlBody(item)])
    .then(([data, isHtml]) => {
      const { url, text } = buildMeetingMessage(data, isHtml, POLYCOM_ENABLED);
      const coercionType = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;

      return new Promise((resolve, reject) => {
        item.body.setSelectedDataAsync(text, { coercionType }, (setResult) => {
          if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(setResult.error);
            return;
          }
          if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            item.location.setAsync(url, () => resolve());
            return;
          }
          resolve();
        });
      });
    })
    .then(() => {
      _showRemoveButton();
    })
    .catch((err) => {
      console.error(err);
    })
    .finally(() => {
      _setButtonIdle();
    });
}

function removeMeetingLinkFromItem() {
  // TEST ONLY — ProConnect auth skipped.
  _setRemoveLoading();

  const item = Office.context.mailbox.item;

  removeMeetingLink(item)
    .then(() => {
      _showAddButton();
    })
    .catch((err) => {
      console.error(err);
    })
    .finally(() => {
      _setRemoveIdle();
    });
}

// ── Init ─────────────────────────────────────────────────────

Office.onReady(async (info) => {
  await initI18n();
  translateUI();

  if (FEEDBACK_FORM) {
    const link = document.getElementById("feedback-link");
    link.href = FEEDBACK_FORM;
    link.style.display = "inline";
  }

  if (info.host === Office.HostType.Outlook) {
    applyAppName();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("btn-generate").onclick = generateMeetingLink;
    document.getElementById("btn-remove").onclick = removeMeetingLinkFromItem;
    _renderPolycomStatus();

    // TEST ONLY — always show the authenticated view, ProConnect bypassed.
    showView("auth"); // this already calls _refreshMeetingButtonState internally
  }
});
