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
    statusElement.innerHTML = "<p style='color: #2b579a;'>מעבד דיווח מאובטח...</p>";
  }

  // מקבלים את ה-ID ומוודאים שהוא בפורמט תקין ל-XML
  let itemId = Office.context.mailbox.item.itemId;
  
  // ניקוי תווים מיוחדים למניעת שגיאת Internal Error
  itemId = itemId.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // פקודת XML לשרת (EWS)
  const ewsRequest = 
    `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                   xmlns:m="http://schemas.microsoft.com/microsoft/exchange/services/2006/messages" 
                   xmlns:t="http://schemas.microsoft.com/microsoft/exchange/services/2006/types" 
                   xmlns:soap="http://schemas.xml-schema.org/soap/envelope/">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
      </soap:Header>
      <soap:Body>
        <m:CreateItem MessageDisposition="SendOnly">
          <m:Items>
            <t:Message>
              <t:Subject>דיווח על מייל חשוד - OFIRSEC</t:Subject>
              <t:Body BodyType="HTML">מצורף מייל שדווח כחשוד על ידי משתמש קצה באמצעות תוסף OFIRSEC.</t:Body>
              <t:Attachments>
                <t:ItemAttachment>
                  <t:Name>Suspicious_Email.eml</t:Name>
                  <t:Item>
                    <t:ItemId Id="${itemId}" />
                  </t:Item>
                </t:ItemAttachment>
              </t:Attachments>
              <t:ToRecipients>
                <t:Mailbox>
                  <t:EmailAddress>Info@ofirsec.co.il</t:EmailAddress>
                </t:Mailbox>
              </t:ToRecipients>
            </t:Message>
          </m:Items>
        </m:CreateItem>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("EWS Success!");
      if (statusElement) {
        statusElement.innerHTML = `
          <div style="text-align:center; margin-top: 20px;">
            <p style="color:green; font-weight:bold; font-size:18px;">✅ הדיווח נשלח בהצלחה!</p>
            <p>תודה על ערנותך, המייל הועבר לטיפול.</p>
          </div>`;
      }
    } else {
      // הדפסה מפורטת ל-Console כדי שנוכל לאבחן במקרה של תקלה
      console.error("EWS Failed Status: " + asyncResult.status);
      console.error("Error Details: ", asyncResult.error);
      
      if (statusElement) {
        statusElement.innerHTML = `
          <div style="color:red; margin-top: 20px;">
            <p>❌ חלה שגיאה פנימית בשליחה.</p>
            <small>קוד שגיאה: ${asyncResult.error.code}</small>
          </div>`;
      }
    }
  });
}