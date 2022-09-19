//Source : git@github.com:menafrancisco/Send-Email-with-Gmail-Draft.git

function getDataAndSendEmails() {
  //Get data and subject placeholder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  let subject = sh.getRange(2, 1).getValue();
  let subjectPH = "";
  let rSubject;
  let cc;
  let body;
  const lastColumn = sh.getLastColumn();
  const [header, ...data] = sh.getRange(3, 1, sh.getLastRow() - 2, lastColumn)
    .getValues();

/*
  //Extract subject placeholder if there is one
  const subjectPH1 = subject.match(/\{\{\w+\}\}/);
  if (subjectPH1 !== null) {
    subjectPH = subjectPH1[0];
  }
  else {
    rSubject = subject;
  }*/

  //Loop thru data, replace body placeholders
  data.forEach((row, r) => {
    body = getDraft(body, subject);

    rSubject = subject;
    
    //Replace body placeholders
    header.forEach((hCol, h) => {

      if (hCol.includes("{{") && hCol !== "{{EMAIL}}" && hCol !== "{{email}}" && hCol.includes("{{") && hCol !== "{{CC}}" && hCol !== "{{cc}}") {

        let rg = new RegExp(hCol, "g");
        console.log(row[h]);
        body = body.replace(rg, row[h]);

        //if (hCol === subjectPH) {
          rSubject = rSubject.replace(rg, row[h]);
        //};
      }

      if (hCol === "{{EMAIL}}" || hCol === "{{email}}") {
        email = row[h];
      };

      if (hCol === "{{CC}}" || hCol === "{{cc}}") {
        cc = row[h];
      };

    });


    var draft=true;
    if(draft){

      //Create draft
      createDraft(body, email, rSubject, cc);

      //Update status on EMAILS sheet
      let rw = r + 4;
      sh.getRange(rw, sh.getLastColumn())
        .setValue("DRAFT CREATED").setBackground('#6fa8dc');
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(1000);
      }

    }else{
      //Send email
      sendEmail(body, email, rSubject);

      //Update status on EMAILS sheet
      let rw = r + 4;
      sh.getRange(rw, sh.getLastColumn())
        .setValue("SENT").setBackground('#93c47d');
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(1000);
      }

    }

  });
}

function getDraft(body, subject) {
  const drafts = GmailApp.getDrafts();

  drafts.forEach((draft) => {
    let draftId = draft.getId();
    let draftById = GmailApp.getDraft(draftId);
    let msg = draftById.getMessage();
    let draftSubject = msg.getSubject();

    if (draftSubject === subject) {
      body = msg.getBody();
    };
  });
  return body;
}

function sendEmail(body, email, rSubject) {
  GmailApp.sendEmail(email,
    rSubject, "",
    { htmlBody: body });
}

function createDraft(body, email, rSubject, cc){
  // Create the email draft
  GmailApp.createDraft(email,
    rSubject, "",
    { htmlBody: body , cc: cc });
}




