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

  // 1. פתיחת חלונית הדיווח - זה החלק שפותח את המייל לשליחה
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: uniqueSubject,
    htmlBody: "שלום צוות אבטחה, אני מדווח על המייל המצורף כחשוד כפישינג.",
    attachments: [{
      type: Office.MailboxEnums.AttachmentType.Item,
      name: "Suspicious_Email",
      itemId: item.itemId
    }]
  });

  // 2. העברה לתיקיית Junk
  // שימוש במחרוזת "junk" עוקף את השגיאה שראית בתמונה
  if (item.moveItemAsync) {
    item.moveItemAsync("junk", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        statusElement.innerHTML = `
          <div style='color:green;'>
            <b>הדיווח נפתח והמייל הועבר ל-Junk!</b><br>
            אל תשכח ללחוץ על 'שלח' בחלון שנפתח.
          </div>`;
      } else {
        console.error("Move failed: " + result.error.message);
        statusElement.innerHTML = "<div style='color:green;'><b>חלון הדיווח נפתח!</b></div>";
      }
    });
  } else {
    // אם בכל זאת הפונקציה לא קיימת (למשל אם העדכון של ה-XML עוד לא נכנס לתוקף מלא)
    statusElement.innerHTML = "<div style='color:green;'><b>חלון הדיווח נפתח!</b></div>";
    console.warn("moveItemAsync is still not available.");
  }
}