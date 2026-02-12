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
  if (statusElement) statusElement.innerHTML = "<p>מכין דיווח...</p>";

  // שימוש ב-createForwardAsync - פונקציה הרבה יותר נתמכת
  Office.context.mailbox.item.createForwardAsync({
    'toRecipients': ['Info@ofirsec.co.il'],
    'htmlBody': 'מצורף דיווח על מייל חשוד.'
  }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const forwardItem = asyncResult.value;
      // פותח את חלון המייל שהוכן
      forwardItem.display();
      if (statusElement) statusElement.innerHTML = "<p style='color:green;'>הדיווח מוכן! לחצי על 'שלח' בחלון שנפתח.</p>";
    } else {
      console.error(asyncResult.error.message);
      if (statusElement) statusElement.innerHTML = "<p style='color:red;'>שגיאה ביצירת הדיווח. נסי שוב.</p>";
    }
  });
}