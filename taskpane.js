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
    statusElement.innerHTML = "<p style='color: #2b579a;'>מבצע דיווח אוטומטי...</p>";
  }

  // בדיקה אם הפונקציה קיימת בגרסה שלך
  if (Office.context.mailbox.item.forwardAsAttachment) {
    Office.context.mailbox.item.forwardAsAttachment(
      ["Info@ofirsec.co.il"],
      {
        subject: "דיווח על מייל חשוד - OFIRSEC",
        htmlBody: "המייל המצורף דווח כחשוד על ידי המשתמש."
      },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          statusElement.innerHTML = "<p style='color:green; font-weight:bold;'>✅ הדיווח נשלח בהצלחה כקובץ מצורף!</p>";
        } else {
          console.error(asyncResult.error);
          statusElement.innerHTML = "<p style='color:red;'>שגיאה בשליחה: " + asyncResult.error.message + "</p>";
        }
      }
    );
  } else {
    // הודעה למקרה שהאאוטלוק ישן מדי
    statusElement.innerHTML = "<p style='color:orange;'>הגרסה של Outlook אינה תומכת בשליחה אוטומטית כקובץ.</p>";
    console.log("forwardAsAttachment is not supported in this Requirement Set.");
  }
}