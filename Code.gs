//Source : https://www.bazroberts.com/2021/11/03/mail-merge-using-draft-emails/

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

  const template= sh.getRange(2, 2).getValue(); 
  
  sh.getRange(2, sh.getLastColumn())
        .setValue("PROCESSING...").setFontColor('#e69138');
      SpreadsheetApp.flush();

  cleanstatus(sh, data);

  //Loop thru data, replace body placeholders
  data.forEach((row, r) => {
    //body = getDraft(body, subject);
    body = template;

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
        .setValue("DRAFT CREATED").setFontColor('#6fa8dc');
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
        .setValue("SENT").setFontColor('#93c47d');
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(1000);
      }
    }
  });


  let flag = true;
  let color='';
  data.forEach((row, r) => {
    //Update status on EMAILS sheet
    let rw = r + 4;
    
    sh.getRange(rw, sh.getLastColumn())
        .setValue("PROCESSED").setFontColor('#000000');
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(1000);
      }
  });  

  sh.getRange(2, sh.getLastColumn())
        .setValue("COMPLETED").setFontColor('#6aa84f');
      SpreadsheetApp.flush();

}

function cleanstatus(sh, data) {
  data.forEach((row, r) => {
    //Update status on EMAILS sheet
    let rw = r + 4;    
    sh.getRange(rw, sh.getLastColumn())
        .setValue("");
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(1000);
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




