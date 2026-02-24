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

  // 1. פתיחת חלונית הדיווח - זה תמיד עובד
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

  // 2. העברה ל-Junk
  // אנחנו בודקים אם הפונקציה קיימת. אם היא לא קיימת, סימן שה-XML 1.5 עוד לא פעיל.
  if (item.moveItemAsync) {
    // השתמשי בערך "junk" באותיות קטנות - זה הכי אמין
    item.moveItemAsync("junk", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        statusElement.innerHTML = "<b style='color:green;'>דווח והועבר ל-Junk בהצלחה!</b>";
      } else {
        // אם זה נכשל, נדפיס למה
        console.log("Move failed: " + result.error.message);
        statusElement.innerHTML = "<b>הדיווח נפתח!</b> (העברה אוטומטית תפעל בקרוב)";
      }
    });
  } else {
    // הודעה זו אומרת שאאוטלוק עדיין לא קרא את ה-XML החדש שלך
    console.warn("The move command is not yet active in your Outlook version.");
    statusElement.innerHTML = "<b>הדיווח נפתח!</b><br><small>שים לב: העברה ל-Junk תהיה זמינה לאחר סיום עדכון המערכת.</small>";
  }
}