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
  if (statusElement) statusElement.innerHTML = "<p style='color: #2b579a;'>מבצע דיווח אוטומטי...</p>";

  // שימוש ב-REST API - הדרך היחידה שעוקפת CORS וחסימות EWS בווב
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const accessToken = result.value;
      const itemId = Office.context.mailbox.item.itemId;
      
      // המרת ה-ID לפורמט תקין ל-REST
      const restId = itemId.replace(/\//g, '-').replace(/\+/g, '_');
      
      // כתובת ה-API של מיקרוסופט
      const url = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + restId + "/forward";

      fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          "Comment": "דיווח אוטומטי על מייל חשוד - OFIRSEC",
          "ToRecipients": [
            { "EmailAddress": { "Address": "Info@ofirsec.co.il" } }
          ]
        })
      })
      .then(response => {
        if (response.ok) {
          statusElement.innerHTML = "<div style='text-align:center; color:green;'><b>✅ הדיווח נשלח אוטומטית!</b></div>";
        } else {
          throw new Error('Server rejected the request');
        }
      })
      .catch(error => {
        console.error("REST Error:", error);
        statusElement.innerHTML = "<p style='color:red;'>שגיאה: השרת חוסם שליחה אוטומטית.</p>";
      });
    } else {
      statusElement.innerHTML = "<p style='color:red;'>שגיאת הרשאה. פנה למנהל המערכת.</p>";
    }
  });
}