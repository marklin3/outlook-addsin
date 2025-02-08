function getUserGists(user, callback) {
  const requestUrl = "https://api.github.com/users/" + user + "/gists";

  $.ajax({
    url: requestUrl,
    dataType: "json",
  })
    .done(function (gists) {
      callback(gists);
    })
    .fail(function (error) {
      callback(null, error);
    });
}

function buildGistList(parent, gists, clickFunc) {
  gists.forEach(function (gist) {
    const listItem = $("<div/>").appendTo(parent);

    const radioItem = $("<input>")
      .addClass("ms-ListItem")
      .addClass("is-selectable")
      .attr("type", "radio")
      .attr("name", "gists")
      .attr("tabindex", 0)
      .val(gist.id)
      .appendTo(listItem);

    const descPrimary = $("<span/>").addClass("ms-ListItem-primaryText").text(gist.description).appendTo(listItem);

    const descSecondary = $("<span/>")
      .addClass("ms-ListItem-secondaryText")
      .text(" - " + buildFileList(gist.files))
      .appendTo(listItem);

    const updated = new Date(gist.updated_at);

    const descTertiary = $("<span/>")
      .addClass("ms-ListItem-tertiaryText")
      .text(" - Last updated " + updated.toLocaleString())
      .appendTo(listItem);

    listItem.on("click", clickFunc);
  });
}

function buildFileList(files) {
  let fileList = "";

  for (let file in files) {
    if (files.hasOwnProperty(file)) {
      if (fileList.length > 0) {
        fileList = fileList + ", ";
      }

      fileList = fileList + files[file].filename + " (" + files[file].language + ")";
    }
  }

  return fileList;
}

function getGist(gistId, callback) {
  const requestUrl = "https://api.github.com/gists/" + gistId;

  $.ajax({
    url: requestUrl,
    dataType: "json",
  })
    .done(function (gist) {
      callback(gist);
    })
    .fail(function (error) {
      callback(null, error);
    });
}

function queryAll(user, callback) {
  // Create a unique callback function name
  const callbackName = `jsonp_callback_${new Date().getTime()}`;

  // Dynamically create a new script element
  const script = document.createElement("script");

  // Add the callback name to the URL as a query parameter
  script.setAttribute(
    "src",
    `https://10.17.62.23:5001/webapi/query.cgi?api=SYNO.API.Info&method=query&version=1&query=all&callback=${callbackName}`
  );

  // Define the global callback function
  window[callbackName] = function (response) {
    // Call the provided callback with the result
    callback(response);

    // Clean up: Remove the callback function and the script tag
    delete window[callbackName];
    document.body.removeChild(script);
  };

  // Handling errors
  script.onerror = function () {
    callback(null, new Error("Failed to load script"));
    // Clean up in case of error
    delete window[callbackName];
    document.body.removeChild(script);
  };

  // Append the script to the document to start the request
  document.body.appendChild(script);
}

function buildBodyContent(gist, callback) {
  // Find the first non-truncated file in the gist
  // and use it.
  for (let filename in gist.files) {
    if (gist.files.hasOwnProperty(filename)) {
      const file = gist.files[filename];
      if (!file.truncated) {
        // We have a winner.
        switch (file.language) {
          case "HTML":
            // Insert as is.
            callback(file.content);
            break;
          case "Markdown":
            // Convert Markdown to HTML.
            const converter = new showdown.Converter();
            const html = converter.makeHtml(file.content);
            callback(html);
            break;
          default:
            // Insert contents as a <code> block.
            let codeBlock = "<pre><code>";
            codeBlock = codeBlock + file.content;
            codeBlock = codeBlock + "</code></pre>";
            callback(codeBlock);
        }
        return;
      }
    }
  }
  callback(null, "No suitable file found in the gist");
}
