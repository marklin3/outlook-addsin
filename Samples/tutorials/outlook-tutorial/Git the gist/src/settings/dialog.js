(function () {
  "use strict";

  // The onReady function must be run each time a new page is loaded.
  Office.onReady(function () {
    $(document).ready(function () {
      // eslint-disable-next-line no-undef
      window.addEventListener("message", function (event) {
        // eslint-disable-next-line no-undef

        if (event.origin === "https://marktest.syno") {
          console.log("Received message: data " + event.data + " origin " + event.origin);

          // Assuming the message data is in event.data
          var data = event.data;

          // Set the content in the email body if Office is available
          // var content = "The meeting link is <a href='https://10.17.62.23:5001/meet/authed/" + data + "'>Join Meeting</a><br>";
          var content =
            "The meeting link is <a href='https://10.17.62.23:5001/meet/authed/" +
            data +
            "' target='_blank'>Join Meeting</a><br>";

          Office.context.ui.messageParent(content);
        }
      });
    });
  });
})();
