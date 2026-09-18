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
    document.