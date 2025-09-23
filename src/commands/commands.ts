/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
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
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/**
 * Opens a dialog when the dialog button is clicked.
 * @param event
 */
function openDialog(event: Office.AddinCommands.Event) {
  const dialogUrl = "https://www.microsoft.com";
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 60, width: 60 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        
        // Handle dialog events
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
          console.log("Message received from dialog:", message);
          dialog.close();
        });
        
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (eventArgs) => {
          console.log("Dialog event:", eventArgs);
          if ('error' in eventArgs && eventArgs.error === 12006) {
            // Dialog was closed by user
            console.log("Dialog was closed by user");
          }
        });
      } else {
        console.error("Failed to open dialog:", result.error);
      }
      
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
  );
}

// Register the function with Office.
Office.actions.associate("action", action);
Office.actions.associate("openDialog", openDialog);
