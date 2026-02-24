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
    htmlBody: "שלום צוות אבטחה, אני מדווח על המייל המצורף כחשוד כפישינג.",
    attachments: [{
      type: Office.MailboxEnums.AttachmentType.Item,
      name: "Suspicious_Email",
      itemId: item.itemId
    }]
  });

  // 2. העברה לתיקיית Junk
  // שימוש ב-"junk" כמחרוזת טקסט חסין לטעויות
  if (item && item.moveItemAsync) {
    item.moveItemAsync("junk", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        statusElement.innerHTML = "<div style='color:green;'><b>הדיווח נפתח והמייל הועבר ל-Junk!</b></div>";
      } else {
        console.error("Move failed: " + result.error.message);
        statusElement.innerHTML = "<div style='color:green;'><b>הדיווח נפתח!</b></div>";
      }
    });
  } else {
    // אם הגעת לכאן, זה אומר שהאאוטלוק עדיין לא "מעוקל" על גרסת המניפסט החדשה (1.5)
    console.warn("moveItemAsync is not yet available on this item.");
    statusElement.innerHTML = "<div style='color:green;'><b>הדיווח נפתח!</b><br><small>(העברה אוטומטית תהיה זמינה לאחר סיום עדכון המערכת)</small></div>";
  }
}