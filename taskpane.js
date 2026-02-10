Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // התאמה ל-ID של הכפתור ב-HTML (שמו שם הוא "run")
    const reportButton = document.getElementById("run");
    if (reportButton) {
      reportButton.onclick = reportEmail;
    }
  }
});

function reportEmail() {
  // התאמה ל-ID של הדיב ב-HTML (שמו שם הוא "status-message")
  const statusElement = document.getElementById("status-message");
  
  if (statusElement) {
    statusElement.innerHTML = "<p style='color: #2b579a;'>מעבד דיווח...</p>";
  }

  // פונקציית השליחה של אופיס
  Office.context.mailbox.item.forwardAsAttachment(
    ["Info@ofirsec.co.il"], 
    {
      asyncContext: null
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // הצלחה - המייל נשלח ברקע
        if (statusElement) {
          statusElement.innerHTML = `
            <div style="text-align:center; margin-top: 20px;">
              <p style="color:green; font-weight:bold; font-size:18px;">✅ הדיווח נשלח בהצלחה!</p>
              <p>תודה על ערנותך. צוות האבטחה יבדוק את המייל.</p>
            </div>
          `;
        }
      } else {
        // שגיאה
        console.error(asyncResult.error.message);
        if (statusElement) {
          statusElement.innerHTML = "<p style='color:red; margin-top: 20px;'>❌ תקלה בשליחה. אנא נסה שוב מאוחר יותר או דווח ידנית.</p>";
        }
      }
    }
  );
}