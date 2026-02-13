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
  if (statusElement) statusElement.innerHTML = "<p style='color: #2b579a;'>שולח דיווח אוטומטי מאובטח...</p>";

  // קבלת טוקן הגישה בזכות הרשאת ReadWriteMailbox
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const accessToken = result.value;
      const itemId = Office.context.mailbox.item.itemId;
      
      // המרת ה-ID לפורמט REST
      const restId = itemId.replace(/\//g, '-').replace(/\+/g, '_');
      const serviceUrl = Office.context.mailbox.restUrl + '/v2.0/me/messages/' + restId + '/forward';

      // בקשת ה-Fetch עוקפת את בעיות ה-CORS שראינו ב-Console שלך
      fetch(serviceUrl, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          "Comment": "דיווח אוטומטי על מייל חשוד - תוסף OFIRSEC",
          "ToRecipients": [
            { "EmailAddress": { "Address": "Info@ofirsec.co.il" } }
          ]
        })
      })
      .then(response => {
        if (response.status === 202) {
          statusElement.innerHTML = "<div style='text-align:center; color:green;'><b>✅ הדיווח נשלח בהצלחה!</b><br>צוות האבטחה עודכן.</div>";
        } else {
          return response.json().then(err => { throw err; });
        }
      })
      .catch(error => {
        console.error("REST Error Details:", error);
        statusElement.innerHTML = "<p style='color:red;'>השרת חוסם שליחה אוטומטית. נא לפנות ל-IT.</p>";
      });
    } else {
      statusElement.innerHTML = "<p style='color:red;'>שגיאת הרשאה בקבלת Token.</p>";
    }
  });
}