Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const reportButton = document.getElementById("run");
    if (reportButton) {
      reportButton.onclick = reportEmail;
    }
  }
});

function reportEmail() {
  const statusElement = document.getElementById("status-message");
  statusElement.innerHTML = "<p style='color: #2b579a;'>מכין את הדיווח...</p>";

  // יצירת חותמת זמן (מילי-שניות) כדי שהנושא יהיה חד-ערכי לחלוטין
  const timestamp = Date.now();
  
  // הנושא החדש שיאפשר ל-Flow ול-SharePoint למצוא את המייל בקלות
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // פתיחת חלון הודעה חדשה עם הנושא הייחודי והמייל המקורי מצורף
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: uniqueSubject,
    htmlBody: "שלום צוות אבטחה,<br><br>אני מדווח על המייל המצורף כחשוד כפישינג.",
    attachments: [
      {
        type: Office.MailboxEnums.AttachmentType.Item,
        name: "Suspicious_Email",
        itemId: Office.context.mailbox.item.itemId
      }
    ]
  });

  statusElement.innerHTML = "<div style='color:green;'><b>חלון הדיווח נפתח!</b><br>אנא לחץ על 'שלח' כדי להשלים את הדיווח.</div>";
}