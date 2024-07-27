function main() {
  distributeDocuments()
}

function distributeDocuments() {

// SECTION 1 - PREPARATION ---------------------------------------------------------------------------------------

  // accessing the folder where all companies' folders will be located
  const targetFolderID = "YOUR ID HERE";
  const targetFolder = DriveApp.getFolderById(targetFolderID);

  // some docs are not generated, but copied - getting access to them
  const pdfs = [
                DriveApp.getFileById("YOUR ID HERE"),
                DriveApp.getFileById("YOUR ID HERE"),
                DriveApp.getFileById("YOUR ID HERE"),
                DriveApp.getFileById("YOUR ID HERE"),
                DriveApp.getFileById("YOUR ID HERE"),
                DriveApp.getFileById("YOUR ID HERE"),
              ] 

  // accessing sheet with responses
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("YOUR SHEET NAME HERE");

  // getting data from original sheet in form of embedded lists
  const values = sheet.getDataRange().getValues();

  // getting IDs of other documents we will move - using functions implemented below
  const idColumns = getIdColumns(values[0])
  const uploadedDocumentsColumns = getUploadedDocumentsColumns(values[1]);

// SECTION 2 - DATA ITERATION --------------------------------------------------------------------------------------------

  // iterating over each row in the sheet, except for headings
  for (let rowNum=1; rowNum < values.length; rowNum++) {
    
    let row = values[rowNum];
    let companyName = row[idColumns[0] - 7];
    Logger.log(companyName);

    // checking 2 times if sending documents is permitted (to prevent accidents)
    if (row[idColumns[0] - 1] != "Так") {
      continue
    }

    if (row[idColumns[0] - 2] != "") {
      Logger.log(row[idColumns[0] - 2])
      continue
    }

    // creating company's personal folder and 4 folders inside it -------------------------- 
    
    let companyFolder = targetFolder.createFolder(companyName + "YOUR FOLDER NAME HERE");
    let uploadedDocumentsFolder = companyFolder.createFolder("YOUR FOLDER NAME HERE"); // an inside folder which we will use in code

    let insideFolders = [
                        companyFolder.createFolder("YOUR FOLDER NAME HERE"), 
                        companyFolder.createFolder("YOUR FOLDER NAME HERE"),
                        companyFolder.createFolder("YOUR FOLDER NAME HERE")
                        ]; // 3 of 4 inside folders are not used in code - so I did not give them a name

    let teamEmails = row[29].split(",");

    // giving access to all folders for all team members ----------------------------------
    for (let i=0; i < teamEmails.length; i++) {
      let email = teamEmails[i].trim();

      companyFolder.addViewer(email);
      
      for (let j=0; j < insideFolders.length; j++) {
        insideFolders[j].addEditor(email);
      }
    }

    // moving all generated documents to company's folder ----------------------------------
    for (let i=0; i < idColumns.length; i++) {
      
      let documentID = row[idColumns[i]]
      let document = DriveApp.getFileById(documentID);
      document.moveTo(companyFolder);

      // giving editor access
      if (document.getName().includes("YOUR FOLDER NAME HERE")) {        

        for (let j=0; j < teamEmails.length; j++) {
          let email = teamEmails[j].trim();
          document.addEditor(email);
        }
      }
    }

    // moving all pdfs' copies to company's folder --------------------------------------------
    for (let i=0; i < pdfs.length; i++) {
      let pdf = pdfs[i].makeCopy();

      pdf.setName(pdf.getName() + " - " + companyName);
      pdf.moveTo(companyFolder);
    }

    // moving all documents uploaded by user to corresponding folder ----------------------------
    for (let i=0; i < uploadedDocumentsColumns.length; i++) {
      
      if (row[uploadedDocumentsColumns[i]] == "") {
        continue
      }  
      
      let documentID = row[uploadedDocumentsColumns[i]].split("?id=")[1];
      let document = DriveApp.getFileById(documentID);

      document.makeCopy().moveTo(uploadedDocumentsFolder);
    }

    // giving access to all folders and documents to communication email -----------------------
    let communicationEmail = row[28];

    companyFolder.addViewer(communicationEmail);
    for (let i=0; i < insideFolders.length; i++) {
      insideFolders[i].addEditor(communicationEmail)
    }

    for (let i=0; i < idColumns.length; i++) {
      let documentID = row[idColumns[i]]
      let document = DriveApp.getFileById(documentID);

      // giving editor access to Appendix №8 - Budget and Financial report
      if (document.getName().includes("YOUR FOLDER NAME HERE")) {        
        document.addEditor(communicationEmail)
      }
    }


    // SECTION 3 - SENDING EMAIL AND ADDING COMPANIES' FOLDER URL TO THE COLUMN ------------------------------

    // adding a url to company's folder in the spreadsheet
    sheet.getRange(rowNum + 1, idColumns[0]-2 + 1).setValue(companyFolder.getUrl());

    let body =  `
    YOUR HTML BODY HERE
    `

    GmailApp.sendEmail(communicationEmail, "YOUR HEADER HERE - " + companyName, "", {htmlBody: body})
  }

}

function getIdColumns(header) {
  let returnColumns = [];

  for (let i=0; i < header.length; i++) {
    if (header[i].includes("Merged Doc ID")){
      returnColumns.push(i);
    }  
  }

  return returnColumns
}

function getUploadedDocumentsColumns(row) {

  resultColumns = [];
  for (let i=0; i < row.length; i++) {
    if (row[i].toString().includes("https://drive.google.com/open?id=")) {
      resultColumns.push(i);
    }
  }

  return resultColumns
}

function copyGranteesFiles() {
  const motherFolderID = "YOUR ID HERE";
  const archiveFolderID = "YOUR ID HERE";

  const motherFolder = DriveApp.getFolderById(motherFolderID);
  const archiveFolder = DriveApp.getFolderById(archiveFolderID); 

  const granteesFolders = motherFolder.getFolders();

  while (granteesFolders.hasNext()) {
  
    let folder = granteesFolders.next();
    let granteesArchieveFolder = archiveFolder.createFolder(folder.getName().split(" - ")[0] + " - .");

    Logger.log(folder.getName().split(" - ")[0]);

    let subfolders = folder.getFolders();

    while (subfolders.hasNext()) {
      let subfolder = subfolders.next()

      if (subfolder.getName().includes(".")) {
        
        files = subfolder.getFiles()

        while (files.hasNext()) {
          let file = files.next().makeCopy();

          file.moveTo(granteesArchieveFolder);
          
        }
 
      }
    }
      
  }
}
