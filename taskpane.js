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
  statusElement.innerHTML = "מדווח ומבודד את המייל...";

  // 1. פתיחת חלונית הדיווח (החלק שחייב להישאר כדי שהמשתמש ילחץ 'שלח')
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: "חשד לפישינג: " + item.subject,
    attachments: [{
      type: Office.MailboxEnums.AttachmentType.Item,
      name: "Suspicious_Email",
      itemId: item.itemId
    }]
  });

  // 2. פקודת "מאחורי הקלעים" להעברת המייל ל-Junk
  // שיטה זו עוקפת את מגבלת הגרסה שנתקלת בה
  const itemId = item.itemId;
  const ewsRequest = 
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
             <t:ItemId Id="${itemId}"/>
           </m:ItemIds>
         </m:MoveItem>
       </soap:Body>
     </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      statusElement.innerHTML = "<b style='color:green;'>הדיווח נפתח והמייל הועבר ל-Junk.</b>";
    } else {
      // אם גם זה נכשל, כנראה שיש חסימה רוחבית ב-IT על תוספים
      console.error(result.error);
      statusElement.innerHTML = "<b>הדיווח נפתח!</b><br>אנא סגור את המייל המקורי.";
    }
  });
}