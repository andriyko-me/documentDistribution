function main() {
  copyGranteesFiles()
}

function copyGranteesFiles() {
  const motherFolderID = "1sGso_SiX7FHUBxizTyYZsnYBaWvFnqB3";
  const archiveFolderID = "1GSIjXBNsQSB-thpZGhkF7sOUw8hrPUVo";

  const motherFolder = DriveApp.getFolderById(motherFolderID);
  const archiveFolder = DriveApp.getFolderById(archiveFolderID); 

  const granteesFolders = motherFolder.getFolders();

  while (granteesFolders.hasNext()) {
  
    let folder = granteesFolders.next();
    let granteesArchieveFolder = archiveFolder.createFolder(folder.getName().split(" - ")[0] + " - Підтверджуючі документи");

    Logger.log(folder.getName().split(" - ")[0]);

    let subfolders = folder.getFolders();

    while (subfolders.hasNext()) {
      let subfolder = subfolders.next()

      if (subfolder.getName().includes("Бюджет")) {
        
        files = subfolder.getFiles()

        while (files.hasNext()) {
          let file = files.next().makeCopy();

          file.moveTo(granteesArchieveFolder);
          
        }
 
      }
    }
      
  }
}

function distributeDocuments() {

// SECTION 1 - PREPARATION ---------------------------------------------------------------------------------------

  // accessing the folder where all companies' folders will be located
  const targetFolderID = "1fQsJz-wD6Clj56JJbv5jHxFB0nOtpwQ7";
  const targetFolder = DriveApp.getFolderById(targetFolderID); // 2. папки грантерів

  // some docs are not generated, but copied - getting access to them
  const pdfs = [
                DriveApp.getFileById("1Q2_kdLz6j1myrzF8WuT13u9E_9CxIUPk"), // Додаток №1
                DriveApp.getFileById("1tgt2KjnnL9oQdwk59SIzQSME9liBXy-r"), // Додаток №3
                DriveApp.getFileById("14rn7McjcD9FDWNxzOQ--yKXHHHTI4D9c"), // Додаток №4
                DriveApp.getFileById("1itel8-7i_RIsumvQpFA135va0Vcarksc"), // Додаток №5
                DriveApp.getFileById("1HhKlaBO9KWsfi63dPOH3uoDDFflAwTtc"), // Додаток №6
                DriveApp.getFileById("1wAphumvsBf55-E83oMaeplZ2bXBJT89o"), // Додаток №7
              ] 

  // accessing sheet with responses
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form Responses 1");

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

    // 2. папки грантоотримувачів --> companies' personal folders --> insideFolders
    
    let companyFolder = targetFolder.createFolder(companyName + " - Грантові документи");
    let uploadedDocumentsFolder = companyFolder.createFolder("Загальні документи грантоотримувача"); // an inside folder which we will use in code

    let insideFolders = [
                        companyFolder.createFolder("Підтверджуючі документи для Бюджету"), 
                        companyFolder.createFolder("Підтверджуючі документи для Фінансового звіту про надходження та використання коштів"),
                        companyFolder.createFolder("Підтверджуючі документи для Звіту про виконання проєкту")
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

      // giving editor access to Appendix №8 - Budget and Financial report
      if (document.getName().includes("Додаток №8")) {        

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

      document.makeCopy().moveTo(uploadedDocumentsFolder); // Загальні документи грантоотримувача
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
      if (document.getName().includes("Додаток №8")) {        
        document.addEditor(communicationEmail)
      }
    }


    // SECTION 3 - SENDING EMAIL AND ADDING COMPANIES' FOLDER URL TO THE COLUMN ------------------------------

    // adding a url to company's folder in the spreadsheet
    sheet.getRange(rowNum + 1, idColumns[0]-2 + 1).setValue(companyFolder.getUrl());

    let body =  `
    Шановні грантоотримувачі,<br><br>
    Дякуємо, що якісно заповнили гугл форму та надали документи/інформацію необхідну для підписання грантового договору. Раді повідомити, що ми з командою згенерували всі необхідні грантові документи та <b>завантажили їх у папку вашого грантового проєкту за посиланням</b>:<br><br>` + companyFolder.getUrl() + `
    <br><br>
    <b>Перевірте, будь ласка, чи немає помилок у всіх згенерованих документах.</b> Особливу увагу зверніть на реквізити та IBAN. <br><br>
    Якщо ви знайдете помилки або у вас виникнуть додаткові запитання, то напишіть їх у відповідь на цей лист.<br><br>
    <b>Якщо всі документи правильні то підпишіть грантовий договір та додатки до нього у системі електронного документообміну Вчасно.</b> Плануємо завтра відправити документи Вам на підпис. <b>Підпишіть договір до 18:00 28.05.24</b><br><br><br>


    У подальшому будемо <b>просити вас ставити запитання саме як відповідь у межах цієї переписки</b> (натиснути кнопку “відповісти” на цей лист, а не створювати новий лист), щоб нам простіше було координувати взаємодію з усіма грантерами.<br><br>


    <b>Ваші наступні кроки: (<u>дедлайн до 31.05.24</u>):</b><br>
    1. Переглянути бюджет, який ви подавали разом з вашою грантовою заявкою, та запланувати на його основі актуальний бюджет вашого грантового проєкту.<br>
    2. Завершити тендери і на основі проведених тендерів у обраних постачальників зібрати пакет рахунків на оплату, які підтверджуватимуть всі заплановані у бюджеті витрати.<br>
    3. Запланувати бюджет на всю суму гранту – 427,290.00 UAH, радимо додати невелике співфінансування з власних коштів. Це на випадок, якщо в процесі реалізації проєкту сума фактичних витрат зменшиться і становитиме менше ніж 10 тис. євро (427,290.00 UAH). У такому разі, під час звітування, ви зменшите заплановані витрати за власні кошти, перерозподіливши їх на витрати за рахунок гранту, і зможете відзвітувати за використання гранту у повному розмірі – 10 тис. євро (427,290.00 UAH).<br>
    4. Якщо категорії витрат або вартість відрізняється в актуальному бюджеті Грантового проєкту від грантової заявки, то необхідно подати запит на зміну бюджету. Використовуйте <a href="https://forms.gle/B5u39gxm6kn1yzgSA">гугл форму</a>.<br>
    5. Перевірити всі рахунки постачальників згідно чекліста, виправити за потреби (див. <a href="https://www.notion.so/yaroslavzhydyk/100-10000-c654deae0dab45ed83e534be1e02e79e">NOTION</a>)<br>
    6. Переглянути відео інструкцію - <a href="https://youtu.be/fOO2paBq2S4">Як заповнювати бюджет</a>.<br>
    <b>7. Заповнити Додаток №8 Бюджет Проєкту на основі рахунків постачальників в гугл таблиці “Додаток №8 Бюджет проєкт та документ Фінансовий звіт”, лист “Бюджет”.<br><br>
    Гугл таблиця знаходиться в папці проєкту посилання на яку ви отримали у цьому листі.</b><br><br>
    <b>8. Завантажити рахунки на оплату, які будуть підтверджувати всі заплановані в бюджеті витрати, в папку “Підтверджуючі документи для Бюджету”. Тендерні документи завантажувати непотрібно, тільки рахунки. Формат PDF. Переіменуйте всі файли так, щоб назва файлу відповідала номеру підкатегорії витрат з Додатка №8 "Бюджет Проєкту".</b><br><br>
    <b>9. Написати імейл про те, що подали на перевірку Додаток №8 Бюджет Проєкту на перевірку <u>до 31 травня 2024 року.</u></b><br><br>

    <a href="https://forms.gle/B5u39gxm6kn1yzgSA">Google Форму "Погодження змін бюджету"</a>, яку ми підготували, просимо заповнити, якщо у вас виникла потреба зробити зміни в бюджеті, <b>окремо просимо написати нам імейл з запитом</b> та вказати що надали всю необхідну інформацію в Google формі "Погодження змін бюджету"<br><br>
    І не забувайте про Покроковий план дій Грантоотримувача та дедлайни. Знайдете їх на платформі <a href="https://www.notion.so/yaroslavzhydyk/100-10000-c654deae0dab45ed83e534be1e02e79e">NOTION</a>. Збережіть це посилання і не втрачайте його, адже там ви зможете знайти відповіді на майже всі ваші запитання.<br><br>
    Спокійного вам дня!<br>
    З повагою,<br>
    Команда проєкту<br><br>

    --<br><br>

    <i>
    Хвиля грантів сприятиме реалізації цілей урядової ініціативи #єРобота під егідою <a href="https://www.me.gov.ua/?lang=uk-UA">Міністерство економіки України</a>.<br><br>

    100 грантів по 10 000 євро кожен надаються у межах  програми міжнародної співпраці «EU4Business: відновлення, конкурентоспроможність та інтернаціоналізація МСП» за фінансування Європейського Союзу та уряду Німеччини. Програма спрямована на підтримку економічної стійкості, відновлення та зростання України, створення кращих умов для розвитку українських малих і середніх підприємств (МСП), а також підтримку інновацій та експорту. Детальніше: www.eu4business.org.ua.<br><br>

    Стратегічний виконавець програми – німецька федеральна компанія <a href="https://www.facebook.com/gizukraine">Deutsche Gesellschaft für Internationale Zusammenarbeit (GIZ) GmbH</a>. Партнер з виконання – громадська організація <a href="https://www.facebook.com/easybusiness.in.ua">EasyBusiness</a> (ГО «Легкий бізнес»).<br><br>

    Зміст листа є виключною відповідальністю EasyBusiness і не обов’язково відображає позицію Європейського Союзу та уряду Німеччини.<br>
    #eu4business, #StandForUkraine, #gizSME
    `

    // GmailApp.sendEmail(communicationEmail, "100 грантів по 10 000 євро - " + companyName, "", {htmlBody: body})
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
