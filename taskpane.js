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
  statusElement.innerHTML = "<p style='color: #2b579a;'>מכין את הדיווח ומנקה את תיבת הדואר...</p>";

  // 1. יצירת חותמת זמן ייחודית
  const timestamp = Date.now();
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // 2. פתיחת חלונית הדיווח (Display New Message Form)
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: uniqueSubject,
    htmlBody: "שלום צוות אבטחה,<br><br>אני מדווח על המייל המצורף כחשוד כפישינג.",
    attachments: [
      {
        type: Office.MailboxEnums.AttachmentType.Item,
        name: "Suspicious_Email",
        itemId: item.itemId
      }
    ]
  });

  // 3. העברת המייל המדווח לתיקיית דואר זבל (Junk)
  // הפעולה הזו קורית ברקע מיד לאחר פתיחת החלונית
  item.moveItemAsync(Office.MailboxEnums.StandardFolder.Junk, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      statusElement.innerHTML = `
        <div style='color:green;'>
          <b>חלון הדיווח נפתח!</b><br>
          המייל המקורי הועבר לתיקיית <b>דואר זבל</b>.<br>
          אנא לחץ על 'שלח' בחלון שנפתח.
        </div>`;
    } else {
      // במקרה של שגיאה בהעברה (למשל אם המייל כבר הועבר), עדיין נציג שהדיווח נפתח
      console.error("Move item failed: " + result.error.message);
      statusElement.innerHTML = "<div style='color:green;'><b>חלון הדיווח נפתח!</b><br>אנא לחץ על 'שלח' כדי להשלים את הדיווח.</div>";
    }
  });
}