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
    statusElement.innerHTML = "<p style='color: #2b579a;'>מבצע דיווח אוטומטי שקט...</p>";
  }

  // זו השיטה הכי חזקה לשליחה אוטומטית בלי EWS ובלי CORS
  // היא שולחת את המייל המקורי כקובץ מצורף (Attachment)
  Office.context.mailbox.item.forwardAsAttachment(
    ["Info@ofirsec.co.il"], // הכתובת שלכם
    {
      subject: "דיווח אוטומטי על מייל חשוד - OFIRSEC",
      htmlBody: "הודעה זו נשלחה אוטומטית מהתוסף. המייל החשוד מצורף כקובץ."
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        if (statusElement) {
          statusElement.innerHTML = `
            <div style="text-align:center; margin-top: 20px;">
              <p style="color:green; font-weight:bold; font-size:18px;">✅ הדיווח נשלח בהצלחה!</p>
              <p>תודה על העירנות. המייל הועבר ל-SOC.</p>
            </div>`;
        }
      } else {
        // אם יש שגיאה, נדפיס אותה כדי להבין מה חסום
        console.error("שגיאה בשליחה אוטומטית:", asyncResult.error);
        statusElement.innerHTML = `
          <div style="color:red; margin-top: 20px;">
            <p>❌ חלה שגיאה בשליחה האוטומטית.</p>
            <p>קוד שגיאה: ${asyncResult.error.code}</p>
          </div>`;
      }
    }
  );
}