/* eslint-disable no-undef */
(function () {
  "use strict";

  let config;
  let settingsDialog;

  // The onReady function must be run each time a new page is loaded.
  Office.onReady();

  // eslint-disable-next-line no-undef
  document.getElementById("connect-button").addEventListener("click", function () {
    // eslint-disable-next-line no-undef
    var nasAddress = document.getElementById("nas-address").value;
    // eslint-disable-next-line no-undef
    var iframe = document.getElementById("meeting-iframe");
    // iframe.src = nasAddress;
    // iframe.src = "https://marktest.syno:443/meet/e7fd0342-e2fe-4935-a75f-176197aaedc4/summary/8";
    // iframe.style.display = "block";
  });

  // eslint-disable-next-line office-addins/no-office-initialize
  Office.initialize = function () {
    jQuery(document).ready(function () {
      config = getConfig();

      // Check if add-in is configured.
      if (config && config.gitHubUserName) {
        // If configured, load the gist list.
        loadGists(config.gitHubUserName);
      } else {
        // Not configured yet.
        $("#not-configured").show();
      }

      // When insert button is selected, build the content
      // and insert into the body.
      $("#connect-button").on("click", function () {
        var nasAddress = document.getElementById("nas-address").value;
        var topic = document.getElementById("meeting-topic").value;


        // // eslint-disable-next-line no-undef
        // var iframe = document.getElementById("meeting-iframe");
        // // iframe.src = nasAddress;
        // iframe.src = "https://marktest.syno:443/meet/e7fd0342-e2fe-4935-a75f-176197aaedc4/summary/8";
        // iframe.style.display = 'block';
        // Display settings dialog.
        // let url = new URI("https://marktest.syno:443/meet/e7fd0342-e2fe-4935-a75f-176197aaedc4/summary/8").toString();
        // let url = new URI("https://10.17.62.23:5001/meet?topic=test&type=webniar").toString();
        let url = new URI("dialog.html").absoluteTo(window.location).toString();
        let data = { address: nasAddress, roomName: topic };
        let queryParams = new URLSearchParams(data).toString();
        let fullUrl = `${url}?${queryParams}`;


        // window.open(url, "_blank");
        // window.addEventListener("message", function (event) {
        //     console.log("Received message: " + event.data);
        // });
        const dialogOptions = { width: 80, height: 120, displayInIframe: true };
        Office.context.ui.displayDialogAsync(fullUrl, dialogOptions, function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            settingsDialog = result.value;
            settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
            settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
          } else {
            console.error("Failed to display dialog: " + result.error.message);
          }
        });
      });
    });
  };

  function getConfig() {
    return Office.context.roamingSettings.get("config");
  }

  function sendQueryAllRequest() {
    const apiUrl = "https://10.17.62.23:5001/webapi/query.cgi?api=SYNO.API.Info&method=query&version=1&query=all";

    // Prepare the data to be sent in the request body
    const requestData = {
      name: "John Doe",
      email: "john.doe@example.com",
    };

    // Send the POST request using fetch
    fetch(apiUrl, {
      method: "GET", // HTTP method
      headers: {
        "Content-Type": "application/json", // Indicate that we're sending JSON
      },
      // body: JSON.stringify(requestData), // Convert the data object to JSON
    })
      .then((response) => response.json()) // Parse the JSON response
      .then((data) => {
        console.log("Success:", data); // Handle the success response
      })
      .catch((error) => {
        console.error("Error:", error); // Handle any errors
      });
  }

  function loadGists(user) {
    $("#error-display").hide();
    $("#not-configured").hide();
    $("#gist-list-container").show();

    getUserGists(user, function (gists, error) {
      if (error) {
      } else {
        $("#gist-list").empty();
        buildGistList($("#gist-list"), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $("#insert-button").removeAttr("disabled");
    $(".ms-ListItem").removeClass("is-selected").removeAttr("checked");
    $(this).children(".ms-ListItem").addClass("is-selected").attr("checked", "checked");
  }

  function showError(error) {
    $("#not-configured").hide();
    $("#gist-list-container").hide();
    $("#error-display").text(error);
    $("#error-display").show();
  }

  function receiveMessage(data) {
    console.log("Received message: " + data.message);

    Office.context.mailbox.item.body.setSelectedDataAsync(
      data.message,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (settingsDialog !== null) {
          settingsDialog.close();
          settingsDialog = null;
        }
      }
    );
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
