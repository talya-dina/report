Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const reportButton = document.getElementById("run");
    if (reportButton) reportButton.onclick = reportEmail;
  }
});

async function reportEmail() {
  const statusElement = document.getElementById("status-message");
  statusElement.innerHTML = "<p style='color: #2b579a;'>מעבד דיווח...</p>";

  // ניסיון א': שליחה אוטומטית שקטה (REST API)
  try {
    const accessToken = await new Promise((resolve, reject) => {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
        if (result.status === "succeeded") resolve(result.value);
        else reject(result.error);
      });
    });

    const itemId = Office.context.mailbox.item.itemId;
    const restId = itemId.replace(/\//g, '-').replace(/\+/g, '_');
    const serviceUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${restId}/forward`;

    const response = await fetch(serviceUrl, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        "Comment": "דיווח אוטומטי - OFIRSEC",
        "ToRecipients": [{ "EmailAddress": { "Address": "Info@ofirsec.co.il" } }]
      })
    });

    if (response.status === 202) {
      statusElement.innerHTML = "<p style='color:green; font-weight:bold;'>✅ הדיווח נשלח אוטומטית!</p>";
      return; // הצלחנו, עוצרים כאן
    }
    throw new Error("Automatic send blocked");

  } catch (error) {
    // ניסיון ב' (Fallback): פתיחת חלון מוכן מראש
    console.log("Starting Fallback: Opening forward form...");
    
    Office.context.mailbox.item.displayForwardForm({
      toRecipients: ["Info@ofirsec.co.il"],
      htmlBody: "מצורף דיווח על מייל חשוד.",
      attachments: [{
        type: Office.MailboxEnums.AttachmentType.Item,
        name: "SuspiciousEmail",
        itemId: Office.context.mailbox.item.itemId
      }]
    });

    statusElement.innerHTML = `
      <div style="color: #856404; background-color: #fff3cd; padding: 10px; border-radius: 5px;">
        <p><b>כמעט סיימנו!</b></p>
        <p>הדיווח הוכן. אנא לחצי על <b>'שלח' (Send)</b> בחלון שנפתח.</p>
      </div>`;
  }
}