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
  
  if (statusElement) {
    statusElement.innerHTML = "<p style='color: #2b579a;'>מכין את הדיווח...</p>";
  }

  // פתיחת חלון העברה (Forward) עם המייל המקורי כקובץ מצורף
  Office.context.mailbox.item.displayForwardForm({
    'toRecipients': ['Info@ofirsec.co.il'],
    'htmlBody': 'שלום צוות אבטחת מידע,<br><br>אני מדווח על המייל המצורף כחשוד.<br><br>בברכה,',
    'attachments': [
      {
        'type': Office.MailboxEnums.AttachmentType.Item,
        'name': 'Suspicious_Email',
        'itemId': Office.context.mailbox.item.itemId
      }
    ]
  });

  // עדכון סטטוס בחלונית לאחר פתיחת החלון
  if (statusElement) {
    statusElement.innerHTML = `
      <div style="text-align:center; margin-top: 20px;">
        <p style="color: #2b579a; font-weight:bold;">הדיווח מוכן!</p>
        <p>חלון מייל חדש נפתח. אנא לחץ על <b>'שלח'</b> כדי להעביר אותו לצוות האבטחה.</p>
      </div>
    `;
  }
}