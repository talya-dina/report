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
    statusElement.innerHTML = "<p style='color: #2b579a;'>שולח דיווח לצוות האבטחה...</p>";
  }

  // קבלת ה-ID של המייל הנוכחי
  const itemId = Office.context.mailbox.item.itemId;

  // פקודת XML לביצוע Forward אוטומטי (עוקף את שגיאה 9020)
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
            <t:ForwardItem>
              <t:ToRecipients>
                <t:Mailbox>
                  <t:EmailAddress>Info@ofirsec.co.il</t:EmailAddress>
                </t:Mailbox>
              </t:ToRecipients>
              <t:ReferenceItemId Id="${itemId.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}" />
              <t:NewBodyContent BodyType="HTML">מצורף דיווח על מייל חשוד שנשלח דרך תוסף OFIRSEC.</t:NewBodyContent>
            </t:ForwardItem>
          </m:Items>
        </m:CreateItem>
      </soap:Body>
    </soap:Envelope>`;

  // ביצוע הבקשה מול השרת
  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      if (statusElement) {
        statusElement.innerHTML = `
          <div style="text-align:center; margin-top: 20px;">
            <p style="color:green; font-weight:bold; font-size:18px;">✅ הדיווח נשלח בהצלחה!</p>
            <p>תודה על ערנותך, צוות האבטחה קיבל את הדיווח.</p>
          </div>`;
      }
    } else {
      console.error("EWS Error: ", asyncResult.error);
      if (statusElement) {
        statusElement.innerHTML = `
          <div style="color:red; margin-top: 20px;">
            <p>❌ תקלה בשליחת הדיווח.</p>
            <small>קוד שגיאה: ${asyncResult.error.code}</small>
          </div>`;
      }
    }
  });
}