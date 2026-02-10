Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("report-button").onclick = reportEmail;
  }
});

function reportEmail() {
  const statusElement = document.getElementById("status");
  statusElement.innerHTML = "<p>מעבד דיווח...</p>";

  // שליחה אוטומטית של המייל הנוכחי כקובץ מצורף
  Office.context.mailbox.item.forwardAsAttachment(
    ["Info@ofirsec.co.il"], // הכתובת שלך
    {
      asyncContext: null
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // הצלחה - המייל נשלח ברקע!
        statusElement.innerHTML = `
          <div style="text-align:center;">
            <p style="color:green; font-weight:bold; font-size:18px;">✅ הדיווח נשלח בהצלחה!</p>
            <p>תודה על ערנותך. צוות האבטחה יבדוק את המייל.</p>
          </div>
        `;
      } else {
        // שגיאה
        console.error(asyncResult.error.message);
        statusElement.innerHTML = "<p style='color:red;'>❌ תקלה בשליחה. אנא נסה שוב מאוחר יותר.</p>";
      }
    }
  );
}