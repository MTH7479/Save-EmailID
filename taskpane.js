/* Copy Email ID - Outlook add-in task pane logic */
(function () {
  "use strict";

  var ids = { rest: "", ews: "", net: "" };

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      setStatus("האדין נטען מחוץ ל-Outlook", true);
      return;
    }

    loadIds();

    document.getElementById("btnRest").onclick = function () {
      copy(ids.rest, "Graph / REST Id");
    };

    document.getElementById("btnEws").onclick = function () {
      copy(ids.ews, "EWS ItemId");
    };

    document.getElementById("btnInternet").onclick = function () {
      copy(ids.net, "Internet Message-Id");
    };

    // Refresh when the user selects another message
    try {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        loadIds
      );
    } catch (e) {
      // not supported in all hosts
    }
  });

  function loadIds() {
    var item = Office.context.mailbox.item;

    if (!item) {
      setStatus("לא נבחר מייל", true);
      return;
    }

    ids.ews = item.itemId || "";
    ids.net = item.internetMessageId || "";
    ids.rest = "";

    if (ids.ews) {
      try {
        ids.rest = Office.context.mailbox.convertToRestId(
          ids.ews,
          Office.MailboxEnums.RestVersion.v2_0
        );
      } catch (e) {
        ids.rest = ids.ews;
      }
    } else if (item.saveAsync) {
      // Compose mode: the message must be saved before it has an id
      item.saveAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          ids.ews = r.value;

          try {
            ids.rest = Office.context.mailbox.convertToRestId(
              r.value,
              Office.MailboxEnums.RestVersion.v2_0
            );
          } catch (e) {
            ids.rest = r.value;
          }

          render();
        }
      });
    }

    render();
  }

  function render() {
    document.getElementById("restId").value = ids.rest;
    document.getElementById("ewsId").value = ids.ews;
    document.getElementById("netId").value = ids.net;
  }

  function copy(text, what) {
    if (!text) {
      setStatus("אין ערך להעתקה (" + what + ")", true);
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(
        function () {
          setStatus("הועתק: " + what);
        },
        function () {
          fallbackCopy(text, what);
        }
      );
    } else {
      fallbackCopy(text, what);
    }
  }

  function fallbackCopy(text, what) {
    var ta = document.createElement("textarea");

    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";

    document.body.appendChild(ta);

    ta.focus();
    ta.select();

    var ok = false;

    try {
      ok = document.execCommand("copy");
    } catch (e) {
      ok = false;
    }

    document.body.removeChild(ta);

    setStatus(
      ok
        ? "הועתק: " + what
        : "ההעתקה נחסמה - סמן ידנית והקש Ctrl+C",
      !ok
    );
  }

  function setStatus(msg, isError) {
    var el = document.getElementById("status");
    el.textContent = msg;
    el.className = isError ? "err" : "";
  }
})();
