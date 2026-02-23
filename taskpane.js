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
  statusElement.innerHTML = "<p style='color: #2b579a;'>מכין את הדיווח ומנקה את התיבה...</p>";

  const timestamp = Date.now();
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // 1. פתיחת חלונית הדיווח
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

  // 2. העברה לתיקיית Junk - שימוש במחרוזת ישירה כדי למנוע את שגיאת ה-Undefined
  if (item.moveItemAsync) {
      item.moveItemAsync("junk", function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          statusElement.innerHTML = `
            <div style='color:green;'>
              <b>חלון הדיווח נפתח!</b><br>
              המייל המקורי הועבר לתיקיית <b>דואר זבל</b>.<br>
              אנא לחץ על 'שלח' בחלון שנפתח.
        </div>`;
        } else {
          console.error("Move item failed: " + result.error.message);
          statusElement.innerHTML = "<div style='color:green;'><b>חלון הדיווח נפתח!</b></div>";
        }
      });
  } else {
      console.warn("moveItemAsync is not supported on this item.");
  }
}