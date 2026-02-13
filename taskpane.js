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

  // פתיחת חלון הודעה חדשה שבו הכל כבר מוכן
  // שיטה זו עובדת בכל ארגון ובכל גרסת אאוטלוק בלי צורך בהגדרות שרת
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: "דיווח על מייל חשוד - OFIRSEC Security",
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