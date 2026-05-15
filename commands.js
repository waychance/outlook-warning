var externalConfirmed = false;

Office.onReady(function () {
  Office.actions.associate("onMessageSend", onItemSend);
});

function onItemSend(event) {
  if (externalConfirmed) {
    externalConfirmed = false;
    event.completed({ allowEvent: true });
    return;
  }

  var item = Office.context.mailbox.item;

  item.to.getAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var recipients = result.value;
      var hasExternal = false;
      var internalDomain = "@sg";

      for (var i = 0; i < recipients.length; i++) {
        if (recipients[i].emailAddress.toLowerCase().indexOf(internalDomain) === -1) {
          hasExternal = true;
          break;
        }
      }

      if (hasExternal) {
        externalConfirmed = true; // Next send will go through
        event.completed({
          allowEvent: false,
          errorMessage: "⚠️ External recipient detected! Click SEND again to confirm."
        });
      } else {
        event.completed({ allowEvent: true });
      }
    }
  });
}
