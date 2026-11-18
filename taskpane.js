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
  statusElement.innerHTML = "<p style='color: #2b579a;'>מכין את הדיווח...</p>";

  const item = Office.context.mailbox.item;
  const timestamp = Date.now();
  const formattedTime = new Date(timestamp).toLocaleString("he-IL");

  // שליפת נושא המייל המקורי וקיצורו עבור שורת הנושא
  const fullSubject = item.subject ? item.subject.trim() : "ללא נושא";
  const shortSubject = fullSubject.length > 80 ? fullSubject.substring(0, 80) + "..." : fullSubject;
  const uniqueSubject = `דיווח על מייל חשוד: ${shortSubject} | OFIRSEC Security (ID: ${timestamp})`;

  // יצירת שם קובץ בטוח עבור הצרופה על בסיס נושא המייל המקורי
  let safeAttachmentName = fullSubject
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!safeAttachmentName) {
    safeAttachmentName = "Suspicious_Email";
  } else if (safeAttachmentName.length > 60) {
    safeAttachmentName = safeAttachmentName.substring(0, 60).trim();
  }

  // טבלה נקייה וקצרה עם נושא המייל וחותמת הזמן בלבד
  const htmlBody = `
    <div dir="rtl" style="font-family: Arial, sans-serif; color: #333333;">
      <p>שלום צוות אבטחה,</p>
      <p>אני מדווח על המייל המצורף כחשוד כפישינג.</p>
      
      <table style="border-collapse: collapse; width: 100%; max-width: 550px; margin-top: 10px; border: 1px solid #d0d7de; font-size: 14px;">
        <tbody>
          <tr style="background-color: #f6f8fa;">
            <td style="padding: 8px 12px; border: 1px solid #d0d7de; font-weight: bold; width: 140px;">נושא המייל:</td>
            <td style="padding: 8px 12px; border: 1px solid #d0d7de;">${escapeHtml(fullSubject)}</td>
          </tr>
          <tr>
            <td style="padding: 8px 12px; border: 1px solid #d0d7de; font-weight: bold;">זמן דיווח (ID):</td>
            <td style="padding: 8px 12px; border: 1px solid #d0d7de;">${formattedTime} <span style="color: #666; font-family: monospace;">(${timestamp})</span></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;

  // פתיחת חלון הודעה חדשה שבו הכל כבר מוכן
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Info@ofirsec.co.il"],
    subject: uniqueSubject,
    htmlBody: htmlBody,
    attachments: [
      {
        type: Office.MailboxEnums.AttachmentType.Item,
        name: safeAttachmentName,
        itemId: item.itemId
      }
    ]
  });

  statusElement.innerHTML = "<div style='color:green;'><b>חלון הדיווח נפתח!</b><br>אנא לחץ על 'שלח' כדי להשלים את הדיווח.</div>";
}

function escapeHtml(text) {
  if (!text) return "";
  const div = document.createElement("div");
  div.innerText = text;
  return div.innerHTML;
}
