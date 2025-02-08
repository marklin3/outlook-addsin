let config;
let btnEvent;

// The onReady function must be run each time a new page is loaded.
Office.onReady();

function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("github-error", {
    type: "errorMessage",
    message: error,
  });
}

let settingsDialog;

function insertDefaultGist(event) {
  config = getConfig();

  // Check if the add-in has been configured.
  if (config && config.defaultGistId) {
    console.log("line 21");
    // Get the default gist content and insert.
    try {
      getGist(config.defaultGistId, function (gist, error) {
        if (gist) {
          buildBodyContent(gist, function (content, error) {
            if (content) {
              Office.context.mailbox.item.body.setSelectedDataAsync(
                content,
                { coercionType: Office.CoercionType.Html },
                function (result) {
                  event.completed();
                }
              );
            } else {
              showError(error);
              event.completed();
            }
          });
        } else {
          showError(error);
          event.completed();
        }
      });
    } catch (err) {
      showError(err);
      event.completed();
    }
  } else {
    console.log("line 51");

    // Save the event object so we can finish up later.
    btnEvent = event;
    // Not configured yet, display settings dialog with
    // warn=1 to display warning.
    const url = new URI("dialog.html?warn=1").absoluteTo(window.location).toString();
    const dialogOptions = { width: 20, height: 40, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Dialog displayed successfully.");
        settingsDialog = result.value;
        settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
        settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
      } else {
        console.error("Failed to display dialog: " + result.error.message);
      }
    });
  }
}

// Register the function.
Office.actions.associate("insertDefaultGist", insertDefaultGist);

function receiveMessage(message) {
  config = JSON.parse(message.message);
  setConfig(config, function (result) {
    settingsDialog.close();
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
  });
}

function dialogClosed(message) {
  settingsDialog = null;
  btnEvent.completed();
  btnEvent = null;
}
