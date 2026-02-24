Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const reportButton = document.getElementById("run");
    if (reportButton) {
      reportButton.onclick = reportEmail;
    }
  }
});

function reportEmail() {
  const item = Office.context.mailbox.item;
  const statusElement = document.getElementById("status-message");
  statusElement.innerHTML = "<p style='color: #2b579a;'>מעבד דיווח ומנקה את התיבה...</p>";

  const timestamp = Date.now();
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // 1. פתיחת חלונית הדיווח
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: uniqueSubject,
    htmlBody: "שלום צוות אבטחה, אני מדווח על מייל חשוד.",
    attachments: [{
      type: Office.MailboxEnums.AttachmentType.Item,
      name: "Suspicious_Email",
      itemId: item.itemId
    }]
  });

  // 2. העברה לתיקיית Junk (עכשיו נתמך בזכות גרסה 1.5 ב-XML)
  item.moveItemAsync(Office.MailboxEnums.StandardFolder.Junk, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      statusElement.innerHTML = "<div style='color:green;'><b>הדיווח נפתח והמייל הועבר ל-Junk!</b></div>";
    } else {
      console.error("Move failed: " + result.error.message);
      statusElement.innerHTML = "<div style='color:green;'><b>חלון הדיווח נפתח!</b></div>";
    }
  });
}