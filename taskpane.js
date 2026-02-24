Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const reportButton = document.getElementById("run");
    if (reportButton) {
      reportButton.onclick = reportEmail;
    }
  }
});

function reportEmail() {
  const item = Office.context.mailbox.item;
  const statusElement = document.getElementById("status-message");
  statusElement.innerHTML = "<p style='color: #2b579a;'>מעבד דיווח ומנקה את התיבה...</p>";

  const timestamp = Date.now();
  const uniqueSubject = `דיווח על מייל חשוד - OFIRSEC Security (ID: ${timestamp})`;

  // 1. פתיחת חלונית הדיווח (זה תמיד עובד)
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: uniqueSubject,
    htmlBody: "שלום צוות אבטחה, אני מדווח על המייל המצורף כחשוד כפישינג.",
    attachments: [{
      type: Office.MailboxEnums.AttachmentType.Item,
      name: "Suspicious_Email",
      itemId: item.itemId
    }]
  });

  // 2. העברה לתיקיית Junk באמצעות EWS (עוקף את מגבלת הגרסה)
  const ewsId = item.itemId;
  const request = 
    `<?xml version="1.0" encoding="utf-8"?>
     <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
       <soap:Body>
         <m:MoveItem>
           <m:ToFolderId>
             <t:DistinguishedFolderId Id="junkemail"/>
           </m:ToFolderId>
           <m:ItemIds>
             <t:ItemId Id="${ewsId}"/>
           </m:ItemIds>
         </m:MoveItem>
       </soap:Body>
     </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      statusElement.innerHTML = `
        <div style='color:green;'>
          <b>הדיווח נפתח והמייל הועבר ל-Junk!</b><br>
          אל תשכח ללחוץ על 'שלח' בחלון הדיווח.
        </div>`;
    } else {
      // אם EWS נכשל, ננסה את השיטה הרגילה כגיבוי אחרון
      if (item.moveItemAsync) {
        item.moveItemAsync("junk", (res) => {
          statusElement.innerHTML = "<div style='color:green;'><b>הדיווח נפתח!</b></div>";
        });
      } else {
        console.error("Move failed via EWS and API 1.5 is not ready.");
        statusElement.innerHTML = "<div style='color:green;'><b>הדיווח נפתח!</b></div>";
      }
    }
  });
}