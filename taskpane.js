(function () {
  "use strict";

  var emailId = "";

  Office.onReady(function (info) {

    if (info.host !== Office.HostType.Outlook) {
      setStatus("האדין נטען מחוץ ל‑Outlook", true);
      return;
    }

    loadEmailId();

    document.getElementById("btnRest").onclick = function () {
      copy(emailId, "EmailId");
    };

    try {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        loadEmailId
      );
    } catch (e) {
      // Not supported in all Outlook hosts
    }
  });

  function loadEmailId() {

    var item = Office.context.mailbox.item;

    if (!item) {
      setStatus("לא נבחר מייל", true);
      return;
    }

    var ewsId = item.itemId || "";
    emailId = "";

    if (ewsId) {

      try {

        emailId = Office.context.mailbox.convertToRestId(
          ewsId,
          Office.MailboxEnums.RestVersion.v2_0
        );

      } catch (e) {

        emailId = ewsId;

      }

      render();

    } else if (item.saveAsync) {

      item.saveAsync(function (result) {

        if (result.status === Office.AsyncResultStatus.Succeeded) {

          try {

            emailId = Office.context.mailbox.convertToRestId(
              result.value,
              Office.MailboxEnums.RestVersion.v2_0
            );

          } catch (e) {

            emailId = result.value;

          }

          render();
        }
      });
    }
  }

  function render() {

    document.getElementById("restId").value = emailId;

  }

  function copy(text, description) {

    if (!text) {
      setStatus("אין ערך להעתקה", true);
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {

      navigator.clipboard.writeText(text).then(
        function () {
          setStatus("הועתק: " + description);
        },
        function () {
          fallbackCopy(text, description);
        }
      );

    } else {

      fallbackCopy(text, description);

    }
  }

  function fallbackCopy(text, description) {

    var ta = document.createElement("textarea");

    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";

    document.body.appendChild(ta);

    ta.focus();
    ta.select();

    var success = false;

    try {
      success = document.execCommand("copy");
    } catch (e) {
      success = false;
    }

    document.body.removeChild(ta);

    setStatus(
      success
        ? "הועתק: " + description
        : "ההעתקה נחסמה - סמן את הטקסט והקש Ctrl+C",
      !success
    );
  }

  function setStatus(message, isError) {

    var status = document.getElementById("status");

    status.textContent = message;
    status.className = isError ? "err" : "";
  }

})();