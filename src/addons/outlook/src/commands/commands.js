/* global Office */
const { APP_NAME, POLYCOM_ENABLED } = require("../common/index");
const { buildMeetingMessage } = require("../common/messageBuilder");
const { applyAppName, getIsHtmlBody } = require("../common/helpers");
const { initI18n, t } = require("../common/i18n");
const { isMeetingAlreadyAdded } = require("../common/meetingDetector");
// TEST ONLY — bypasses ProConnect auth and real room creation. Not for merge.
const { TEST_ROOM_DATA } = require("../common/testFixture");

Office.onReady(async function (info) {

  await initI18n()

  if (info.host === Office.HostType.Outlook) {
    applyAppName();
  }
});

function notify(message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("meetNotif", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message,
    persistent: false,
    icon: "Icon.16x16",
  });
}

function insertMeetingLink(event) {
  const item = Office.context.mailbox.item;

  isMeetingAlreadyAdded(item)
    .then((alreadyAdded) => {
      if (alreadyAdded) {
        notify(t("meeting.already_added", { app_name: APP_NAME }));
        event.completed();
        return;
      }
      return _doInsertMeetingLink(event);
    })
    .catch((err) => {
      notify(t("meeting.error.details", { message: err.message }));
      event.completed();
    });
}

function _doInsertMeetingLink(event) {
  const item = Office.context.mailbox.item;

  Promise.all([Promise.resolve(TEST_ROOM_DATA), getIsHtmlBody(item)])
    .then(([data, isHtml]) => {
      const { url, text } = buildMeetingMessage(data, isHtml, POLYCOM_ENABLED);
      const coercionType = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;

      return new Promise((resolve, reject) => {
        item.body.setSelectedDataAsync(text, { coercionType }, (setResult) => {
          if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
            notify(t("meeting.error.details", { message: setResult.error.message }));
            resolve();
            return;
          }

          if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
            notify(t("meeting.link_inserted"));
            resolve();
            return;
          }

          item.location.setAsync(url, (locationResult) => {
            if (locationResult.status !== Office.AsyncResultStatus.Succeeded) {
              notify(t("meeting.error.details", { message: locationResult.error.message }));
            } else {
              notify(t("meeting.link_inserted"));
            }
            resolve();
          });
        });
      });
    })
    .catch((err) => {
      notify(`Erreur : ${err.message}`);
    })
    .finally(() => {
      event.completed();
    });
}

function generateMeetingLink(event) {
  // TEST ONLY — ProConnect auth skipped, goes straight to insertion.
  insertMeetingLink(event);
}

Office.actions.associate("generateMeetingLinkFromCalendar", generateMeetingLink);
Office.actions.associate("generateMeetingLinkFromMail", generateMeetingLink);
