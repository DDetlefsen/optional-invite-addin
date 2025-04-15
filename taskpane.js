Office.onReady(() => {
  document.getElementById("convertBtn").onclick = async () => {
    try {
      const item = Office.context.mailbox.item;
      if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
        console.log("Not a meeting item.");
        return;
      }
      await item.requiredAttendees.getAsync(async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const required = result.value;
          await item.requiredAttendees.setAsync([], () => {
            item.optionalAttendees.setAsync(required);
          });
        } else {
          console.error("Failed to get attendees.");
        }
      });
    } catch (e) {
      console.error(e);
    }
  };
});
