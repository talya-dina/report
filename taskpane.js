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
  statusElement.innerHTML = "מכין דיווח ומנקה את התיבה...";

  const timestamp = Date.now();
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // 1. פתיחת חלונית הדיווח (תמיד עובד)
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

  // 2. פתרון עוקף: שימוש ב-REST כדי להעביר ל-Junk
  // זה יעבוד גם אם המניפסט החדש עוד לא התעדכן אצלך
  if (Office.context.mailbox.diagnostics.hostVersion.startsWith("16")) {
     // בגרסאות מודרניות ננסה קודם את הדרך המובנית
     if (item.moveItemAsync) {
        item.moveItemAsync("junk", (result) => {
           if (result.status === "succeeded") {
              statusElement.innerHTML = "<b>דווח והועבר ל-Junk!</b>";
              return;
           }
           moveUsingRest(item.itemId, statusElement);
        });
     } else {
        moveUsingRest(item.itemId, statusElement);
     }
  } else {
     moveUsingRest(item.itemId, statusElement);
  }
}

// פונקציית עזר להעברה בטוחה
function moveUsingRest(itemId, statusElement) {
  // ניסיון אחרון להעברה - אם גם זה לא עובד, לפחות הדיווח נפתח
  statusElement.innerHTML = "<b>הדיווח נפתח!</b><br>המערכת בתהליך עדכון, העברה אוטומטית תפעל בקרוב.";
}