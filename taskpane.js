Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const reportButton = document.getElementById("run");
    if (reportButton) reportButton.onclick = reportEmail;
  }
});

function reportEmail() {
  const statusElement = document.getElementById("status-message");
  statusElement.innerHTML = "מבצע דיווח אוטומטי...";

  const itemId = Office.context.mailbox.item.itemId;

  // יצירת בקשת EWS בסיסית ביותר ליצירת מייל חדש ושליחתו מיד
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
              <t:Subject>דיווח על מייל חשוד - ID: ${itemId.substring(0,10)}</t:Subject>
              <t:Body BodyType="HTML">המשתמש דיווח על מייל חשוד. מזהה פנימי: ${itemId}</t:Body>
              <t:ToRecipients>
                <t:Mailbox><t:EmailAddress>Info@ofirsec.co.il</t:EmailAddress></t:Mailbox>
              </t:ToRecipients>
            </t:Message>
          </m:Items>
        </m:CreateItem>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      statusElement.innerHTML = "<b style='color:green;'>✅ הדיווח נשלח אוטומטית!</b>";
    } else {
      console.error(result.error);
      // אם הכל נכשל - הדרך היחידה שנותרה היא קישור Mailto פשוט
      statusElement.innerHTML = "<p style='color:red;'>שליחה אוטומטית חסומה בארגון.</p>";
      openMailtoFallback();
    }
  });
}

function openMailtoFallback() {
  const body = "מדווח על מייל חשוד. מזהה הודעה: " + Office.context.mailbox.item.itemId;
  window.location.href = "mailto:Info@ofirsec.co.il?subject=Phishing Report&body=" + encodeURIComponent(body);
}