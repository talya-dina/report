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
    statusElement.innerHTML = "<p style='color: #2b579a;'>שולח דיווח אוטומטי סופי...</p>";
  }

  // מקבלים את ה-ID ומנקים אותו לפורמט בטוח ל-XML
  const itemId = Office.context.mailbox.item.itemId;
  const safeItemId = itemId.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

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
              <t:ReferenceItemId Id="${safeItemId}" />
              <t:NewBodyContent BodyType="HTML">מצורף דיווח אוטומטי על מייל חשוד (OFIRSEC).</t:NewBodyContent>
            </t:ForwardItem>
          </m:Items>
        </m:CreateItem>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      statusElement.innerHTML = "<div style='text-align:center;'><p style='color:green; font-weight:bold; font-size:18px;'>✅ נשלח בהצלחה!</p><p>הדיווח התקבל בצוות האבטחה.</p></div>";
    } else {
      // אם גם זה נכשל, אנחנו נציג את השגיאה המדויקת מהשרת
      console.error("Full Error Object:", asyncResult.error);
      statusElement.innerHTML = `
        <div style="color:red;">
          <p>❌ השרת חסם את השליחה האוטומטית.</p>
          <small>קוד: ${asyncResult.error.code}</small>
        </div>`;
    }
  });
}