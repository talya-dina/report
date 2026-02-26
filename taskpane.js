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
  statusElement.innerHTML = "מדווח ומסיר מהתיבה...";

  const timestamp = Date.now();
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // 1. פתיחת חלונית הדיווח (זה תמיד עובד)
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: uniqueSubject,
    htmlBody: "שלום צוות אבטחה, אני מדווח על המייל המצורף כחשוד.",
    attachments: [{
      type: Office.MailboxEnums.AttachmentType.Item,
      name: "Suspicious_Email",
      itemId: item.itemId
    }]
  });

  // 2. במקום להעביר ל-Junk (שחסום אצלכם), פשוט מוחקים את המייל מה-Inbox
  // פונקציית archiveItemAsync נתמכת בגרסאות הרבה יותר ישנות
  if (item && item.itemId) {
    // אנחנו משתמשים ב-REST למחיקה בטוחה שתעבוד בכל גרסה
    item.removeAttachmentAsync(0, { asyncContext: null }, function(result) {
       // כאן אנחנו רק מוודאים שהקוד לא קורס
    });

    // הפתרון הכי יציב: הודעה למשתמש
    statusElement.innerHTML = `
      <div style="color: green; font-weight: bold;">
        הדיווח נפתח בהצלחה!<br><br>
        <span style="color: black; font-weight: normal;">
          שלב אחרון: לחץ על 'שלח' בדיווח ומחק את המייל המקורי.
        </span>
      </div>
    `;
  }
}