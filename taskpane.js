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
  if (statusElement) statusElement.innerHTML = "<p style='color: #2b579a;'>מבצע דיווח אוטומטי מאובטח...</p>";

  // שלב 1: קבלת טוקן (כרטיס כניסה) מהשרת
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const accessToken = result.value;
      const itemId = Office.context.mailbox.item.itemId;
      
      // שלב 2: המרת ה-ID לפורמט שמתאים ל-REST
      const restId = itemId.replace(/\//g, '-').replace(/\+/g, '_');
      const serviceUrl = Office.context.mailbox.restUrl + '/v2.0/me/messages/' + restId + '/forward';

      // שלב 3: שליחת המייל ישירות דרך ה-API של מיקרוסופט
      const xhr = new XMLHttpRequest();
      xhr.open('POST', serviceUrl);
      xhr.setRequestHeader('Content-Type', 'application/json');
      xhr.setRequestHeader('Authorization', 'Bearer ' + accessToken);

      xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
          if (xhr.status === 202) {
            statusElement.innerHTML = "<div style='text-align:center; color:green;'><b>✅ הדיווח נשלח אוטומטית!</b><br>תודה על העירנות.</div>";
          } else {
            console.error("REST Error:", xhr.responseText);
            // אם ה-REST חסום, ננסה את השיטה האחרונה - פתיחת חלון
            statusElement.innerHTML = "<p style='color:red;'>השרת חוסם שליחה אוטומטית. פותח חלון דיווח...</p>";
            fallbackToForwardForm();
          }
        }
      };

      const body = {
        "Comment": "דיווח על מייל חשוד - OFIRSEC",
        "ToRecipients": [
          { "EmailAddress": { "Address": "Info@ofirsec.co.il" } }
        ]
      };

      xhr.send(JSON.stringify(body));
    } else {
      statusElement.innerHTML = "<p style='color:red;'>שגיאת הרשאה. פותח חלון דיווח...</p>";
      fallbackToForwardForm();
    }
  });
}

// שיטת גיבוי למקרה שהארגון חסם הכל
function fallbackToForwardForm() {
  Office.context.mailbox.item.displayForwardForm({
    'toRecipients': ['Info@ofirsec.co.il'],
    'htmlBody': 'מצורף דיווח על מייל חשוד.',
    'attachments': [{
      'type': Office.MailboxEnums.AttachmentType.Item,
      'name': 'Original_Email',
      'itemId': Office.context.mailbox.item.itemId
    }]
  });
}