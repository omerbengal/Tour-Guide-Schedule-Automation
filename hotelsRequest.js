//export files to --> https://drive.google.com/drive/folders/184FFPzGxZsp9pKmDbkuISVcZevzr3vfD

//https://docs.google.com/spreadsheets/d/1UqKNNSSgzev9TJ010AAMImRebQj548hqLuxmfxEvmag/edit#gid=2099109969
const exportHotelsSpreadsheet = SpreadsheetApp.openById("1UqKNNSSgzev9TJ010AAMImRebQj548hqLuxmfxEvmag");
const sheet = exportHotelsSpreadsheet.getSheets()[0];
const ui = SpreadsheetApp.getUi();


function exportToPDF() {
    if (sheet.getRange("HotelName").isBlank() || sheet.getRange("StartDate").isBlank() || sheet.getRange("EndDate").isBlank() || sheet.getRange("RoomsTable").isBlank()) {
        ui.alert("יש למלא שם מלון, תאריך תחילה, תאריך סיום ופירוט חדרים");
    } else {
        if (sheet.getRange("EndDate").getValue() < sheet.getRange("StartDate").getValue()) {
            ui.alert("תאריך סיום לא יכול להיות לפני תאריך תחילה");
        } else {
            if (ui.alert("?האם להריץ", ui.ButtonSet.YES_NO) == ui.Button.YES) {
                sheet.setName(sheet.getRange("HotelName").getDisplayValue() + " - " + sheet.getRange("StartDate").getDisplayValue());
                var exportFile = DriveApp.createFile(exportHotelsSpreadsheet.getBlob()).setName(sheet.getName().toString());
                exportFile.moveTo(DriveApp.getFoldersByName("Hotels exported PDF's").next());
                exportFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
                sendGmail(ui.prompt("?לאיזה אימייל לשלוח את בקשת הצעת המחיר").getResponseText(), exportFile);
                sheet.setName("-");
                sheet.getRange("HotelName").clearContent();
                sheet.getRange("StartDate").clearContent();
                sheet.getRange("EndDate").clearContent();
                sheet.getRange("Israelis").clearContent();
                sheet.getRange("Tourists").clearContent();
                sheet.getRange("ClientName").clearContent();
                sheet.getRange("RoomsTable").clearContent();
            }
        }
    }
}

function sendGmail(toWho, exportFile) {
    //https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(String,String,String,Object)

    var body = ",שלום רב<br>אשמח לקבל הצעת מחיר סוכן עבור הפרטים הנמצאים בקובץ המצורף<br><br>!תודה<br><br>";
    var signature = "​Hava Voliovich<br>​HavaYa Tours Israel<br>​+972-522-555539";
    var html = "<div style='text-align:right;direction:ltr;font-size:16px'>" + body + signature + "</div>";
    MailApp.sendEmail(toWho, "בקשת לקבלת הצעת מחיר סוכן", body + signature, { name: "Havaya Tours Israel", htmlBody: html, attachments: exportFile.getAs(MimeType.PDF) });
}

function openUrl() {
    var html = '<h1></h1><script>window.onload = function(){google.script.run.withSuccessHandler(function(url){window.open(url,"_blank");google.script.host.close();}).getUrl();}</script>';
    SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(html), ".......פותח פורמט אימייל");
}

function getUrl() {
    var subject = "בקשת לקבלת הצעת מחיר סוכן";
    var body = "היי, אשמח לקבל הצעת מחיר סוכן עבור הפרטים הנמצאים בקובץ המצורף" + '\n' + exportFileUrl;
    return "https://mail.google.com/mail/?view=cm&fs=1&to=" + to + "&su=" + subject + "&body=" + body + "&bcc=havaya5@gmail.com"
}