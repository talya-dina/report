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
    statusElement.innerHTML = "<p style='color: #2b579a;'>מעבד דיווח אוטומטי...</p>";
  }

  const itemId = Office.context.mailbox.item.itemId;

  // פקודת XML לשרת לשליחת המייל כקובץ מצורף
  const ewsRequest = 
    `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                   xmlns:m="http://schemas.microsoft.com/microsoft/exchange/services/2006/messages" 
                   xmlns:t="http://schemas.microsoft.com/microsoft/exchange/services/2006/types" 
                   xmlns:soap="http://schemas.xml-schema.org/soap/envelope/">
      <soap:Header><t:RequestServerVersion Version="Exchange2013" /></soap:Header>
      <soap:Body>
        <m:CreateItem MessageDisposition="SendOnly">
          <m:Items>
            <t:Message>
              <t:Subject>דיווח על מייל חשוד - OFIRSEC</t:Subject>
              <t:Body BodyType="HTML">הודעה זו נוצרה אוטומטית על ידי תוסף הדיווח של OFIRSEC. המייל החשוד מצורף כקובץ.</t:Body>
              <t:Attachments>
                <t:ItemAttachment>
                  <t:Name>Suspicious_Email.eml</t:Name>
                  <t:Item><t:ItemId Id="${itemId}" /></t:Item>
                </t:ItemAttachment>
              </t:Attachments>
              <t:ToRecipients>
                <t:Mailbox><t:EmailAddress>Info@ofirsec.co.il</t:EmailAddress></t:Mailbox>
              </t:ToRecipients>
            </t:Message>
          </m:Items>
        </m:CreateItem>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      if (statusElement) {
        statusElement.innerHTML = `
          <div style="text-align:center; margin-top: 20px;">
            <p style="color:green; font-weight:bold; font-size:18px;">✅ הדיווח נשלח בהצלחה!</p>
            <p>המייל הועבר לצוות האבטחה לבדיקה. תודה!</p>
          </div>`;
      }
    } else {
      console.error(asyncResult.error.message);
      if (statusElement) {
        statusElement.innerHTML = "<p style='color:red;'>שגיאה בשליחה. ייתכן שאין הרשאות מתאימות במניפסט.</p>";
      }
    }
  });
}