/**
 * Gmail Smart Email Tracker + History Archiver (ALL-IN-ONE)
 */

/* ========================= MENU ========================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Gmail Tracker")
    .addItem("Setup Tracker Sheets", "setupEmailTracker")
    .addSeparator()
    .addItem("Send All Pending Emails", "sendPendingEmails")
    .addItem("Check Replies", "checkEmailReplies")
    .addItem("Check Seen", "checkEmailSeen")
    .addItem("Store Email History", "storeEmailThreadHistory")
    .addToUi();
}

/* ========================= SETUP ========================= */
function setupEmailTracker() {
  const ss = SpreadsheetApp.getActive();

  let tracker = ss.getSheetByName("Email_Tracker");
  if (!tracker) tracker = ss.insertSheet("Email_Tracker");

  tracker.clear();
  tracker.getRange("A1:L1").setValues([[
    "Name","Email","CC Emails","Subject","Message",
    "Status","Reply From","Reply Message",
    "Sent Date","Reply Date","Thread ID","Seen"
  ]]);
  tracker.setFrozenRows(1);

  let history = ss.getSheetByName("Email_History");
  if (!history) history = ss.insertSheet("Email_History");

  history.clear();
  history.getRange("A1:H1").setValues([[
    "Thread ID","Contact Name","Contact Email",
    "Message Type","Sender","Sender Email",
    "Date/Time","Message Body"
  ]]);
  history.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("Tracker setup complete");
}

/* ========================= RICH TEXT → HTML ========================= */
function richTextToHtml(richText) {
  if (!richText) return "";
  let html = "";

  richText.getRuns().forEach(run => {
    let text = run.getText()
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/\n/g,"<br>");

    const style = run.getTextStyle();
    const link = run.getLinkUrl(); //

    let css = "";
    if (style.isBold()) css += "font-weight:bold;";
    if (style.isItalic()) css += "font-style:italic;";
    if (style.isUnderline()) css += "text-decoration:underline;";
    if (style.getFontSize()) css += `font-size:${style.getFontSize()}px;`;
    if (style.getForegroundColor()) css += `color:${style.getForegroundColor()};`;

    let span = css ? `<span style="${css}">${text}</span>` : text;

    // WRAP LINK PROPERLY
    if (link) {
      span = `<a href="${link}" target="_blank" style="text-decoration:none;">${span}</a>`;
    }

    html += span;
  });

  return html;
}

/* ========================= SEND EMAILS ========================= */
function sendPendingEmails() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Email_Tracker");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  /* ---------- SIGNATURE ---------- */
  const signatureHtml = `
    <div style="font-family:Arial, sans-serif; font-size:13px; color:#888888;">
      <p style="margin:0 0 6px 0;">
        With Best Regards,<br>
        [your name]<br> 
        Managing Director{yor postions}
      </p>

      <img src="Your logo"
           width="140" height="140"
           style="display:block; margin:10px 0;">

      <p style="margin:6px 0; font-size:16px; font-weight:bold;">
        {Name of company}
      </p>

      <p style="margin:2px 0;"> 
        {office address}
      </p>

      <p style="margin:4px 0;">
        <a href="Your website link">www.shreedestinations.com</a>
      </p>

      <p style="margin:2px 0;">
        Office: 022 31837151 / 47485171<br>
        Mobile: +91 9892236461<br>
        <a href="{your g-mail}">
          {your g-mail}
        </a>
      </p>

      <p style="margin-top:10px; font-size:15px; font-style:italic;">
        {company title}
      </p>
    </div>
  `;

  let sent = 0;

  for (let r = 2; r <= lastRow; r++) {
    const name = sheet.getRange(r,1).getValue();
    const email = sheet.getRange(r,2).getValue();
    const ccEmails = sheet.getRange(r,3).getValue();
    const subject = sheet.getRange(r,4).getValue();
    const status = sheet.getRange(r,6).getValue();
    if (status !== "Pending" || !email) continue;

    const richText = sheet.getRange(r,5).getRichTextValue();
    const bodyHtml = richTextToHtml(richText);

    const finalHtml = `
      <p>Dear ${name},</p>
      ${bodyHtml}<br><br>${signatureHtml}
    `;

    GmailApp.sendEmail(email, subject, "", {
      htmlBody: finalHtml,
      cc: ccEmails
    });

    Utilities.sleep(500);

    const threads = GmailApp.search(`to:${email} subject:"${subject}" newer_than:1d`);
    if (threads.length) sheet.getRange(r,11).setValue(threads[0].getId());

    sheet.getRange(r,6).setValue("Sent");
    sheet.getRange(r,9).setValue(new Date());
    sent++;
  }

  SpreadsheetApp.getUi().alert(`${sent} emails sent`);
}

/* ========================= CHECK REPLIES ========================= */
function checkEmailReplies() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Email_Tracker");
  const lastRow = sheet.getLastRow();
  let found = 0;

  for (let r = 2; r <= lastRow; r++) {
    const email = sheet.getRange(r,2).getValue();
    const status = sheet.getRange(r,6).getValue();
    if (!email || status === "Replied") continue;

    const threads = GmailApp.search(`from:${email} newer_than:30d`);
    if (!threads.length) continue;

    const msg = threads[0].getMessages().pop();
    sheet.getRange(r,6).setValue("Replied");
    sheet.getRange(r,7).setValue(msg.getFrom());
    sheet.getRange(r,8).setValue(msg.getPlainBody().slice(0,500));
    sheet.getRange(r,10).setValue(new Date());
    found++;
  }

  SpreadsheetApp.getUi().alert(`${found} replies updated`);
}

/* ========================= CHECK SEEN ========================= */
function checkEmailSeen() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Email_Tracker");
  const lastRow = sheet.getLastRow();
  let seen = 0;

  for (let r = 2; r <= lastRow; r++) {
    const threadId = sheet.getRange(r,11).getValue();
    const status = sheet.getRange(r,6).getValue();
    if (!threadId || status !== "Sent") continue;

    try {
      const thread = GmailApp.getThreadById(threadId);
      const msgs = thread.getMessages();
      if (!msgs[msgs.length - 1].isUnread()) {
        sheet.getRange(r,12).setValue("👁️ Seen");
        seen++;
      }
    } catch(e){}
  }

  SpreadsheetApp.getUi().alert(`${seen} emails seen`);
}

/* ========================= STORE HISTORY ========================= */
function storeEmailThreadHistory() {
  const ss = SpreadsheetApp.getActive();
  const tracker = ss.getSheetByName("Email_Tracker");
  const history = ss.getSheetByName("Email_History");

  const data = tracker.getRange(2,1,tracker.getLastRow()-1,12).getValues();
  let saved = 0;

  data.forEach(row => {
    const name = row[0];
    const email = row[1];
    const threadId = row[10];
    if (!threadId) return;

    try {
      const thread = GmailApp.getThreadById(threadId);
      thread.getMessages().forEach(msg => {
        const sender = msg.getFrom();
        const senderEmail = sender.match(/<(.+)>/)?.[1] || sender;
        const type = senderEmail === Session.getActiveUser().getEmail() ? "Sent" : "Reply";

        history.appendRow([
          threadId,name,email,type,
          sender,senderEmail,
          msg.getDate(),
          msg.getPlainBody().slice(0,1000)
        ]);
        saved++;
      });
    } catch(e){}
  });

  SpreadsheetApp.getUi().alert(`${saved} messages archived`);
}
