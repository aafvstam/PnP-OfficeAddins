import { ensureStateInitialized, SetRuntimeVisibleHelper } from "../../utilities/office-apis-helpers";

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
  ensureStateInitialized(true);
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

export function btnOpenTaskpane(event: Office.AddinCommands.Event) {
  SetRuntimeVisibleHelper(true);
  g.state.isTaskpaneOpen = true;
  event.completed();
}

export function btnCloseTaskpane(event: Office.AddinCommands.Event) {
  SetRuntimeVisibleHelper(false);
  g.state.isTaskpaneOpen = false;
  event.completed();
}

export function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.btnopentaskpane = btnOpenTaskpane;
g.btnclosetaskpane = btnCloseTaskpane;
g.action = action;
