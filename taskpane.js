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
  statusElement.innerHTML = "מכין דיווח...";

  const timestamp = Date.now();
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // 1. פתיחת חלונית הדיווח (זה החלק החשוב ביותר)
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

  // 2. ניסיון להוציא את המייל מה-Inbox (ארכיון/מחיקה)
  if (item.archiveItemAsync) {
    item.archiveItemAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        statusElement.innerHTML = "<b style='color:green;'>הדיווח נפתח והמייל הוסר מהתיבה!</b>";
      } else {
        // אם הארכיון נכשל, ננחה את המשתמש למחוק
        showManualDeleteMessage(statusElement);
      }
    });
  } else {
    // אם גם הפונקציה הזו לא זמינה בגרסה שלך
    showManualDeleteMessage(statusElement);
  }
}

function showManualDeleteMessage(element) {
  element.innerHTML = `
    <div style="color: #2b579a; text-align: right; direction: rtl;">
      <b style="color: green;">הדיווח נפתח בהצלחה!</b><br><br>
      1. לחץ על <b>'שלח'</b> בחלון החדש.<br>
      2. כעת ניתן <b>למחוק</b> את המייל המקורי.
    </div>
  `;
}