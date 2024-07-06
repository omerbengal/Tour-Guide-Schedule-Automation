SpreadsheetApp.flush();

/*
Automation Files: https://drive.google.com/drive/folders/1ZCQI9IbdI3SRRZvgT6UA-ub0KvJB73Tn
Exported PDF's: https://drive.google.com/drive/folders/1iMw_Gd_AE1Lz3--MbQr-6ASElvohMAbb
workflow: https://docs.google.com/document/d/1olGDOTNUkvmmq2cPvPgBapdOPJFxnI7j9G1PDgXQiY4/edit
*/

//old - https://docs.google.com/spreadsheets/d/1c-MkR-gomEfEGcAg54qdN6yydRoULFJAPA4vmu7wnmU/edit#gid=1546604348
//new - https://docs.google.com/spreadsheets/d/1o3yopEyqpGbcPIkKX3M-vjSlXtefkd69wU5kD2L3PBc/edit#gid=2088109306
const hayavaSpreadsheet = SpreadsheetApp.openById("1o3yopEyqpGbcPIkKX3M-vjSlXtefkd69wU5kD2L3PBc");
const sheets = hayavaSpreadsheet.getSheets();
const mainList = hayavaSpreadsheet.getSheetByName("תגובות לטופס 1");
const quarterSheet = hayavaSpreadsheet.getSheetByName("רבעון");

//https://docs.google.com/spreadsheets/d/1UqKNNSSgzev9TJ010AAMImRebQj548hqLuxmfxEvmag/edit#gid=2099109969
//const exportSpreadsheet = SpreadsheetApp.openById("1UqKNNSSgzev9TJ010AAMImRebQj548hqLuxmfxEvmag");
const exportHotelsSpreadsheet = SpreadsheetApp.openById("1UqKNNSSgzev9TJ010AAMImRebQj548hqLuxmfxEvmag");

const calendar = CalendarApp.getCalendarById("havaya5@gmail.com");
const holidyCalendar = CalendarApp.getCalendarById("iw.jewish#holiday@group.v.calendar.google.com");

//https://docs.google.com/forms/d/1OBGW2vxpwAh915jRgdNrJEL77wPIcxc2DIhONCra0ew/edit#response=ACYDBNivplopctse_xCv4r2XsTgCRWlPHwch-NZO01HPAB0NgTevrdV3ISjjOgpj1Wty6GY
const responses = FormApp.openById('1OBGW2vxpwAh915jRgdNrJEL77wPIcxc2DIhONCra0ew').getResponses();

var lastRowValues;
var sheet;
var sheetName;

var startDate;
var endDate;
var endDatePlusOne;
var name;

var lastDayOfTheMonth = ["3101", "2802", "3103", "3004", "3105", "3006", "3107", "3108", "3009", "3110", "3011", "3112"];

//https://docs.google.com/spreadsheets/d/101u0c7JeIy2iXA6bwYcpBKPLiJu7k20JQ4t4sVmeVok/edit#gid=0
const deletedSpreadsheet = SpreadsheetApp.openById("101u0c7JeIy2iXA6bwYcpBKPLiJu7k20JQ4t4sVmeVok");

// new Date() format: https://stackoverflow.com/questions/33160422/how-to-create-a-specific-date-in-google-script

function onSubmit() {

    SpreadsheetApp.flush();

    formatMainListLastRow();

    lastRowValues = mainList.getRange("A" + mainList.getLastRow() + ":Q" + mainList.getLastRow()).getDisplayValues();
    // 0 - Timestamp, 1 - Contact Name, 2 - Company Name, 3 - E-mail, 4 - Phone Number, 5 - Referrer, 6 - Start Date, 7 - End Date, 8 - Start Time, 9 - End Time,
    //10 - Pick-Up Location, 11 - Required Services, 12 - Where To?, 13 - Passengars #, 14 - Language, 15 - Passengars' Age, 16 - Special Requests.

    startDate = new Date(lastRowValues[0][6].split("/")[2], (parseInt(lastRowValues[0][6].split("/")[1]) - 1), lastRowValues[0][6].split("/")[0], 00, 00, 00);


    if (lastDayOfTheMonth.includes(lastRowValues[0][7].split("/")[0] + lastRowValues[0][7].split("/")[1])) { //אם תאריך סיום הוא סוף חודש

        endDate = new Date(lastRowValues[0][7].split("/")[2], (parseInt(lastRowValues[0][7].split("/")[1]) - 1), lastRowValues[0][7].split("/")[0], 00, 00, 00);
        endDatePlusOne = new Date(lastRowValues[0][7].split("/")[2], lastRowValues[0][7].split("/")[1], 1, 00, 00, 00);

    } else {

        endDate = new Date(lastRowValues[0][7].split("/")[2], (parseInt(lastRowValues[0][7].split("/")[1]) - 1), lastRowValues[0][7].split("/")[0], 00, 00, 00);
        endDatePlusOne = new Date(lastRowValues[0][7].split("/")[2], (parseInt(lastRowValues[0][7].split("/")[1]) - 1), (parseInt(lastRowValues[0][7].split("/")[0]) + 1), 00, 00, 00);

    }


    if (lastRowValues[0][2] == "") {
        name = lastRowValues[0][1];
    } else {
        name = lastRowValues[0][2];
    }

    sheetName = name + " - " + lastRowValues[0][6];

    createSheetAndInsertData();

    collisions();

    sheetDesign();

    createTentativeCalendarEvent();

    addToQuarterSheet();

    MailApp.sendEmail("havaya5@gmail.com,omerbengal7+havayaNotifications@gmail.com", sheetName, "!נכנסה תגובת טופס חדשה");
}

function formatMainListLastRow() {

    var lastRow = mainList.getLastRow();

    mainList.getRange("A" + lastRow + ":Q" + lastRow).setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Assistant").setFontSize(12);

    mainList.getRange("A" + mainList.getLastRow()).setNumberFormat("dd/MM/yyyy hh:mm:ss");
    mainList.getRange("I" + mainList.getLastRow() + ":J" + mainList.getLastRow()).setNumberFormat("hh:mm");
    mainList.getRange("G" + mainList.getLastRow() + ":H" + mainList.getLastRow()).setNumberFormat("dd/MM/yyyy");

    mainList.getRange("L" + lastRow).setValue(mainList.getRange("L" + lastRow).getDisplayValue().replace("הדרכה / Guidance", "Guidance").replace("הסעות / Transportation", "Transportation").replace("מלונות / Hotels", "Hotels").replace("ארוחות / Meals", "Meals"));

    mainList.getRange("O" + lastRow).setValue(mainList.getRange("O" + lastRow).getDisplayValue().replace("עברית / Hebrew", "Hebrew").replace("אנגלית / English", "English").replace("צרפתית / French", "French").replace("ספרדית / Spanish", "Spanish"));

    mainList.getRange("P" + lastRow).setValue(mainList.getRange("P" + lastRow).getDisplayValue().replace("תינוקות / Toddlers", "Toddlers").replace("(2-12) ילדים / Children", "Children (2-12)").replace("מתבגרים / Teens", "Teens").replace("מבוגרים / Adults", "Adults").replace("אזרחים ותיקים / Senior", "Senior"));

}

function createSheetAndInsertData() {

    SpreadsheetApp.flush();

    lastRowValues = mainList.getRange("A" + mainList.getLastRow() + ":Q" + mainList.getLastRow()).getDisplayValues();

    sheet = hayavaSpreadsheet.insertSheet(sheetName, 2);

    lastRowValues[0].forEach(function (i) {
        lastRowValues.push([i]);
    });
    lastRowValues.splice(0, 1);

    sheet.getRange("B5").setNumberFormat("@");
    sheet.getRange("B12").setNumberFormat("@");
    mainList.getRange("A1:Q1").copyTo(sheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
    sheet.getRange("B1:B17").setValues(lastRowValues);

}

function collisions() {

    var startTime = "";
    var endTime = "";

    calendar.getEvents(startDate, endDatePlusOne).forEach(function (event, i) {
        if (/[\u0590-\u05FF]/.test(event.getTitle().toString())) {
            sheet.getRange("A" + (i + 21)).setValue(event.getTitle().toString()).setTextDirection(SpreadsheetApp.TextDirection.RIGHT_TO_LEFT);
        } else {
            sheet.getRange("A" + (i + 21)).setValue(event.getTitle().toString()).setTextDirection(SpreadsheetApp.TextDirection.LEFT_TO_RIGHT);
        }

        startTime = event.getStartTime();
        endTime = event.getEndTime();

        if (event.isAllDayEvent()) {
            sheet.getRange("B" + (i + 21)).setValue(Utilities.formatDate(new Date(startTime), Session.getScriptTimeZone(), "dd/MM/yyyy") + " <---> " + Utilities.formatDate(new Date(endTime), Session.getScriptTimeZone(), "dd/MM/yyyy"));
        } else {
            //https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html - hour as 24 format ("kk")
            sheet.getRange("B" + (i + 21)).setValue(Utilities.formatDate(new Date(startTime), Session.getScriptTimeZone(), "dd/MM/yyyy, kk:mm") + " <---> " + Utilities.formatDate(new Date(endTime), Session.getScriptTimeZone(), "dd/MM/yyyy, kk:mm"));
        }
    });

    if (sheet.getLastRow() == 17) {
        sheet.getRange("A21").setValue("אין אירועים מתנגשים");
        sheet.getRange("A21:B21").merge().setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Assistant").setFontSize(12).setFontWeight("bold").setBackground("#ea9999");
    } else {
        sheet.getRange("A21:B" + sheet.getLastRow()).sort({ column: 2, ascending: true });
    }
}

function sheetDesign() {

    var designLastRow = sheet.getLastRow();
    sheet.setRightToLeft(true);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 400);
    sheet.setColumnWidth(3, 75);

    sheet.getRange("C7").setValue(Math.floor((endDatePlusOne - startDate) / (24 * 60 * 60 * 1000))).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue(sheet.getRange("C7").getDisplayValue());
    sheet.getRange("C7:C8").merge();

    sheet.getRange("A20").setValue("אירועים מתנגשים").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange("A20:B20").merge().setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Assistant").setFontSize(14).setFontWeight("bold").setTextDirection(SpreadsheetApp.TextDirection.RIGHT_TO_LEFT).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sheet.getRange("A1:B17").setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Assistant").setFontSize(12).setFontWeight("bold");
    sheet.getRange("C7:C8").setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Assistant").setFontSize(12).setFontWeight("bold");

    sheet.getRange("A21:B" + designLastRow).setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontFamily("Assistant").setFontSize(12).setFontWeight("bold").setBackground("#ea9999");


    sheet.getRange("A1:B17").setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange("C7:C8").setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange("A21:B" + designLastRow).setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange("A1:C" + designLastRow).protect().setDescription("ranges one must not change").setWarningOnly(true);

}

function createTentativeCalendarEvent() {

    calendar.createAllDayEvent(
        name,
        startDate,
        endDatePlusOne,
        { description: "קישור לגיליון: " + hayavaSpreadsheet.getUrl() + "#gid=" + sheet.getSheetId().toString() }

    ).setColor("5").removeAllReminders();
}

function exportSheet() {
    exportSpreadsheet.getSheets()[0].setName("1");
    sheet.copyTo(exportSpreadsheet).setName(sheet.getName());
    exportSpreadsheet.deleteSheet(exportSpreadsheet.getSheets()[0]);
    exportSpreadsheet.getSheets()[0].getRange("B7").setNumberFormat("dd/MM/yyyy");
    exportSpreadsheet.getSheets()[0].getRange("B8").setNumberFormat("dd/MM/yyyy");
    exportSpreadsheet.getSheets()[0].getRange("B9").setNumberFormat("hh:mm");
    exportSpreadsheet.getSheets()[0].getRange("B10").setNumberFormat("hh:mm");
    SpreadsheetApp.flush();
    var exportFile = DriveApp.createFile(exportSpreadsheet.getBlob()).setName(exportSpreadsheet.getSheets()[0].getName());
    exportFile.moveTo(DriveApp.getFoldersByName("Exported PDF's").next());
    //MailApp.sendEmail("omerbengal7@gmail.com", "export :)", "ניסיון PDF",{attachments: exportFile.getAs(MimeType.PDF)});
}

function addToQuarterSheet() {
    SpreadsheetApp.flush();
    lastRowValues = mainList.getRange("A" + mainList.getLastRow() + ":Q" + mainList.getLastRow()).getDisplayValues();
    var rowToInsert = quarterSheet.getLastRow() + 1;
    var requestedServices = lastRowValues[0][11];

    quarterSheet.getRange("A" + rowToInsert).setValue(name);
    quarterSheet.getRange("B" + rowToInsert).setValue(Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "MM/dd/yyyy")).setNumberFormat("dd/MM/yyyy");
    quarterSheet.getRange("C" + rowToInsert).setValue(Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "MM/dd/yyyy")).setNumberFormat("dd/MM/yyyy");
    quarterSheet.getRange("D" + rowToInsert).setValue(lastRowValues[0][12]);

    if (requestedServices.includes("Guidance")) {
        quarterSheet.getRange("F" + rowToInsert).insertCheckboxes();
    }

    if (requestedServices.includes("Transportation")) {
        quarterSheet.getRange("G" + rowToInsert).insertCheckboxes();
    }

    if (requestedServices.includes("Hotels")) {
        quarterSheet.getRange("H" + rowToInsert).insertCheckboxes();
    }

    if (requestedServices.includes("Meals")) {
        quarterSheet.getRange("I" + rowToInsert).insertCheckboxes();
    }

    //"Other" (in requested services)
    requestedServices = requestedServices.replace("Guidance, ", "").replace("Transportation, ", "").replace("Hotels, ", "").replace("Meals, ", "").replace("Guidance", "").replace("Transportation", "").replace("Hotels", "").replace("Meals", "");
    if (String.prototype.concat(...new Set(requestedServices)) != "") {
        quarterSheet.getRange("K" + rowToInsert).setValue(requestedServices);
    }

    //הוספת צ'קבוקס ביטול
    quarterSheet.getRange("L" + rowToInsert).insertCheckboxes();

    var link = hayavaSpreadsheet.getUrl() + "#gid=" + hayavaSpreadsheet.getSheetByName(sheetName).getSheetId().toString();

    //לינק
    quarterSheet.getRange("A" + rowToInsert).setValue("=HYPERLINK(" + '"' + link + '"' + "," + '"' + name.replace('"', '""" &"').replace("'", '" & "' + "'" + '" & "') + '"' + ")");

    //מיון
    quarterSheet.getRange("A2:L" + quarterSheet.getLastRow()).sort({ column: 2, ascending: true });

    //גבולות
    quarterSheet.getRange("A1:L" + quarterSheet.getLastRow()).setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

}

function firstSync() {

    SpreadsheetApp.flush();

    var lastRow = hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getLastRow();

    if (responses.length == lastRow - 1 && responses.length == hayavaSpreadsheet.getSheets().length - 2) {
        hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF1").setValue("מספר התגובות הוא: " + responses.length).setBackground("#bfffa2");
    } else {
        hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF1").setValue("תגובות ברשימה: " + (lastRow - 1) + ", גיליונות: " + (hayavaSpreadsheet.getSheets().length - 2) + ", תגובות בטופס: " + responses.length).setBackground("#ff6666");
    }

    var key = "";
    var sheetExist = false;
    var formResponseExist = false;

    if (hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("B" + lastRow).getValue() == "עברית") {
        key = hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("C" + lastRow).getValue() + " - " + hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("G" + lastRow).getDisplayValue();
    } else {
        key = hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("Q" + counter).getValue() + " - " + hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("U" + lastRow).getDisplayValue();
    }

    //בדיקת המפתח בגיליונות
    if (hayavaSpreadsheet.getSheetByName(key)) {
        sheetExist = true;
    }

    //בדיקת המפתח בתגובות הטופס
    responses.forEach(function (response) {
        if (key == response.getItemResponses()[1].getResponse() + " - " + response.getItemResponses()[5].getResponse().toString().split("-")[2] + "/" + response.getItemResponses()[5].getResponse().toString().split("-")[1] + "/" + response.getItemResponses()[5].getResponse().toString().split("-")[0]) {
            formResponseExist = true;
        }
    });

    if (sheetExist && formResponseExist) {
        hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + lastRow).setValue("גיליון - V , טופס - V").setBackground("#bfffa2").setHorizontalAlignment("center").setVerticalAlignment("middle");
    } else if (!sheetExist && formResponseExist) {
        hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + lastRow).setValue("גיליון - X , טופס - V").setBackground("#ffaa50").setHorizontalAlignment("center").setVerticalAlignment("middle");
    } else if (sheetExist && !formResponseExist) {
        hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + lastRow).setValue("גיליון - V , טופס - X").setBackground("#ffaa50").setHorizontalAlignment("center").setVerticalAlignment("middle");
    } else {
        hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + lastRow).setValue("גיליון - X , טופס - X").setBackground("#ff6666").setHorizontalAlignment("center").setVerticalAlignment("middle");
    }


}

function amountSync() {

    var ui = SpreadsheetApp.getUi();
    var shouldStart = ui.alert("האם להתחיל בדיקה?", ui.ButtonSet.YES_NO);

    if (shouldStart == ui.Button.YES) {

        SpreadsheetApp.flush();

        const lastRow = hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getLastRow();

        if (responses.length == lastRow - 1 && responses.length == hayavaSpreadsheet.getSheets().length - 2) {
            hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF1").setValue("מספר התגובות הוא: " + responses.length).setBackground("#bfffa2");
        } else {
            hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF1").setValue("תגובות ברשימה: " + (lastRow - 1) + ", גיליונות: " + (hayavaSpreadsheet.getSheets().length - 2) + ", תגובות בטופס: " + responses.length).setBackground("#ff6666");
        }

        var key = "";
        var sheetExist = false;
        var formResponseExist = false;

        for (var counter = 2; counter <= lastRow; counter++) {

            if (hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("B" + counter).getValue() == "עברית") {
                key = hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("C" + counter).getValue() + " - " + hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("G" + counter).getDisplayValue();
            } else {
                key = hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("Q" + counter).getValue() + " - " + hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("U" + counter).getDisplayValue();
            }

            //בדיקת המפתח בגיליונות
            if (hayavaSpreadsheet.getSheetByName(key)) {
                sheetExist = true;
            }

            //בדיקת המפתח בתגובות הטופס
            responses.forEach(function (response) {
                if (key == response.getItemResponses()[1].getResponse() + " - " + response.getItemResponses()[5].getResponse().toString().split("-")[2] + "/" + response.getItemResponses()[5].getResponse().toString().split("-")[1] + "/" + response.getItemResponses()[5].getResponse().toString().split("-")[0]) {
                    formResponseExist = true;
                }
            });

            if (sheetExist && formResponseExist) {
                hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + counter).setValue("גיליון - V , טופס - V").setBackground("#bfffa2").setHorizontalAlignment("center").setVerticalAlignment("middle");
            } else if (!sheetExist && formResponseExist) {
                hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + counter).setValue("גיליון - X , טופס - V").setBackground("#ffaa50").setHorizontalAlignment("center").setVerticalAlignment("middle");
            } else if (sheetExist && !formResponseExist) {
                hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + counter).setValue("גיליון - V , טופס - X").setBackground("#ffaa50").setHorizontalAlignment("center").setVerticalAlignment("middle");
            } else {
                hayavaSpreadsheet.getSheetByName("תגובות לטופס 1").getRange("AF" + counter).setValue("גיליון - X , טופס - X").setBackground("#ff6666").setHorizontalAlignment("center").setVerticalAlignment("middle");
            }

            sheetExist = false;
            formResponseExist = false;

        }

        SpreadsheetApp.flush();
        ui.alert("הבדיקה הסתיימה!");

    }

}

function adminSheet() {

    var ui = SpreadsheetApp.getUi();
    var shouldStart = ui.alert("האם להתחיל בדיקה?", ui.ButtonSet.YES_NO);

    if (shouldStart == ui.Button.YES) {

        SpreadsheetApp.flush();

        var adminSheetLastRow = hayavaSpreadsheet.getSheetByName("ניהולי").getLastRow();
        hayavaSpreadsheet.getSheetByName("ניהולי").getRange("A7:U" + adminSheetLastRow + 1).clearContent();

        var sheetsShirshurim = {};

        hayavaSpreadsheet.getSheets().forEach(function (sheet, i) {
            if (sheet.getName() != "תגובות לטופס 1" && sheet.getName() != "ניהולי") {
                sheetsShirshurim[hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B2").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B3").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B4").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B5").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B6").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B7").getDisplayValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B8").getDisplayValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B9").getDisplayValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B10").getDisplayValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B11").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B12").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B13").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B14").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B15").getValue() + "|" + hayavaSpreadsheet.getSheetByName(sheet.getName()).getRange("B16").getValue()] = i;
            }
        });


        var listShirshurim = {};
        for (var i = 2; i <= mainList.getLastRow(); i++) {

            listShirshurim[mainList.getRange("B" + i).getValue() + "|" + mainList.getRange("C" + i).getValue() + "|" + mainList.getRange("D" + i).getValue() + "|" + mainList.getRange("E" + i).getValue() + "|" + mainList.getRange("F" + i).getValue() + "|" + mainList.getRange("G" + i).getDisplayValue() + "|" + mainList.getRange("H" + i).getDisplayValue() + "|" + mainList.getRange("I" + i).getDisplayValue() + "|" + mainList.getRange("J" + i).getDisplayValue() + "|" + mainList.getRange("K" + i).getValue() + "|" + mainList.getRange("L" + i).getValue() + "|" + mainList.getRange("M" + i).getValue() + "|" + mainList.getRange("N" + i).getValue() + "|" + mainList.getRange("O" + i).getValue() + "|" + mainList.getRange("P" + i).getValue()] = i;

        }

        var formShirshurim = {};
        var tempConcat = "";

        responses.forEach(function (response, j) {
            response.getItemResponses().forEach(function (itemResponse, i) {
                if (i == 5 || i == 6) {
                    tempConcat += itemResponse.getResponse().toString().split("-")[2] + "/" + itemResponse.getResponse().toString().split("-")[1] + "/" + itemResponse.getResponse().toString().split("-")[0] + "|";
                } else if (i == 10 || i == 13) {
                    itemResponse.getResponse().toString().split(",").forEach(function (item, i) {
                        if (i == itemResponse.getResponse().toString().split(",").length - 1) {
                            tempConcat += item + "|";
                        } else {
                            tempConcat += item + ", ";
                        }
                    });
                } else {
                    if (i == response.getItemResponses().length - 1) {
                        tempConcat += itemResponse.getResponse();
                    } else {
                        tempConcat += itemResponse.getResponse() + "|";
                    }
                }
            })
            formShirshurim[tempConcat] = j + 1;
            tempConcat = "";
        });


        var tempRow = 7;

        for (let k in listShirshurim) {
            if (!(k in sheetsShirshurim)) {
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("A" + tempRow).setValue(listShirshurim[k]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("B" + tempRow).setValue(k.split("|")[1]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("C" + tempRow).setValue(k.split("|")[5]);
                tempRow++;
            }
        }

        tempRow = 7;

        for (let k in listShirshurim) {
            if (!(k in formShirshurim)) {
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("E" + tempRow).setValue(listShirshurim[k]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("F" + tempRow).setValue(k.split("|")[1]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("G" + tempRow).setValue(k.split("|")[5]);
                tempRow++;
            }
        }

        tempRow = 7;

        for (let k in sheetsShirshurim) {
            if (!(k in listShirshurim)) {
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("I" + tempRow).setValue(sheetsShirshurim[k] + 1);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("J" + tempRow).setValue(hayavaSpreadsheet.getSheets()[sheetsShirshurim[k]].getName());
                tempRow++;
            }
        }

        tempRow = 7;

        for (let k in sheetsShirshurim) {
            if (!(k in formShirshurim)) {
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("L" + tempRow).setValue(sheetsShirshurim[k] + 1);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("M" + tempRow).setValue(hayavaSpreadsheet.getSheets()[sheetsShirshurim[k]].getName());
                tempRow++;
            }
        }

        tempRow = 7;

        for (let k in formShirshurim) {
            if (!(k in listShirshurim)) {
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("O" + tempRow).setValue(formShirshurim[k]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("P" + tempRow).setValue(k.split("|")[1]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("Q" + tempRow).setValue(k.split("|")[5]);
                tempRow++;
            }
        }

        tempRow = 7;

        for (let k in formShirshurim) {
            if (!(k in sheetsShirshurim)) {
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("S" + tempRow).setValue(formShirshurim[k]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("T" + tempRow).setValue(k.split("|")[1]);
                hayavaSpreadsheet.getSheetByName("ניהולי").getRange("U" + tempRow).setValue(k.split("|")[5]);
                tempRow++;
            }
        }

        adminSheetLastRow = hayavaSpreadsheet.getSheetByName("ניהולי").getLastRow();
        hayavaSpreadsheet.getSheetByName("ניהולי").getRange("A7:U" + adminSheetLastRow).setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");

        SpreadsheetApp.flush();
        ui.alert("הבדיקה הסתיימה!");

    }

}

function createSheetAndInsertData_BEFORE_ARRAY() {

    SpreadsheetApp.flush();
    var lastRow = mainList.getLastRow();

    if (mainList.getRange("C" + lastRow).isBlank()) {
        sheet = hayavaSpreadsheet.insertSheet(mainList.getRange("B" + lastRow).getDisplayValue() + " - " + mainList.getRange("G" + lastRow).getDisplayValue(), 2);
    } else {
        sheet = hayavaSpreadsheet.insertSheet(mainList.getRange("C" + lastRow).getDisplayValue() + " - " + mainList.getRange("G" + lastRow).getDisplayValue(), 2);
    }

    mainList.getRange("A1:Q1").copyTo(sheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
    mainList.getRange("A" + lastRow + ":Q" + lastRow).copyTo(sheet.getRange("B1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);

    //להכניס לכל גיליון בשדות צ׳קבוקס (למיניהם) רק את התשובות באנגלית

    var tempConcat = "";
    sheet.getRange("B12").getDisplayValue().split(", ").forEach(
        function (item, i) {
            item.split(" / ").forEach(
                function (insideItem, j) {
                    if (i != 0 && j == 1) {
                        tempConcat += ", " + insideItem.toString();
                    } else if (i == 0 && j == 1) {
                        tempConcat += insideItem.toString();
                    }
                }
            );
        }
    );
    sheet.getRange("B12").setValue(tempConcat);
    tempConcat = "";

    sheet.getRange("B15").getDisplayValue().split(", ").forEach(
        function (item, i) {
            item.split(" / ").forEach(
                function (insideItem, j) {
                    if (i != 0 && j == 1) {
                        tempConcat += ", " + insideItem.toString();
                    } else if (i == 0 && j == 1) {
                        tempConcat += insideItem.toString();
                    }
                }
            );
        }
    );
    sheet.getRange("B15").setValue(tempConcat);
    tempConcat = "";

    sheet.getRange("B16").getDisplayValue().split(", ").forEach(
        function (item, i) {
            item.split(" / ").forEach(
                function (insideItem, j) {
                    if (i != 0 && j == 1) {
                        tempConcat += ", " + insideItem.toString();
                    } else if (i == 0 && j == 1) {
                        tempConcat += insideItem.toString();
                    }
                }
            );
        }
    );
    sheet.getRange("B16").setValue(tempConcat);
    tempConcat = "";

}

function createSheetAndInsertDataOLD() {

    SpreadsheetApp.flush();
    var lastRow = mainList.getLastRow();


    if (mainList.getRange("B" + lastRow).getValue() == "עברית") {
        sheet = hayavaSpreadsheet.insertSheet(mainList.getRange("C" + lastRow).getValue() + " - " + mainList.getRange("G" + lastRow).getDisplayValue(), 2);

        mainList.getRange("A1:P1").copyTo(sheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
        mainList.getRange("A" + lastRow + ":P" + lastRow).copyTo(sheet.getRange("B1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);

    } else {
        sheet = hayavaSpreadsheet.insertSheet(mainList.getRange("Q" + lastRow).getValue() + " - " + mainList.getRange("U" + lastRow).getDisplayValue(), 2);

        mainList.getRange("A1:P1").copyTo(sheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
        mainList.getRange("A" + lastRow + ":B" + lastRow).copyTo(sheet.getRange("B1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
        mainList.getRange("Q" + lastRow + ":AD" + lastRow).copyTo(sheet.getRange("B3"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);

    }
}

function sheetInsertDataOLD() {

    SpreadsheetApp.flush();
    const lastRow = mainList.getLastRow();

    if (mainList.getRange("B" + lastRow).getValue() == "עברית") {
        sheet = hayavaSpreadsheet.insertSheet(mainList.getRange("C" + lastRow).getValue() + " - " + mainList.getRange("G" + lastRow).getDisplayValue(), 2);

        mainList.getRange("A1:P1").copyTo(sheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);

        mainList.getRange("A" + lastRow + ":P" + lastRow).copyTo(sheet.getRange("B1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);

    } else {
        sheet = hayavaSpreadsheet.insertSheet(mainList.getRange("Q" + lastRow).getValue() + " - " + mainList.getRange("U" + lastRow).getDisplayValue(), 2);

        mainList.getRange("A1:P1").copyTo(sheet.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);

        mainList.getRange("A" + lastRow + ":B" + lastRow).copyTo(sheet.getRange("B1"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
        mainList.getRange("Q" + lastRow + ":AD" + lastRow).copyTo(sheet.getRange("B3"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);

    }

}

function exportSheetOLD() {
    var sheetName = "ענב צרף - 19/09/2022";
    var ss = hayavaSpreadsheet;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetName() !== sheetName) {
            sheets[i].hideSheet()
        }
    }
    var newFile = DriveApp.createFile(ss.getBlob());
    for (var i = 0; i < sheets.length; i++) {
        sheets[i].showSheet()
    }
    MailApp.sendEmail("omerbengal7@gmail.com", "hi", "ניסיון PDF", { attachments: newFile.getAs(MimeType.PDF) })
}

function deleteResponse() {

    const ui = SpreadsheetApp.getUi();

    var rowToDelete = null;
    var name = "";
    var date = "";
    var problemString = "Problem with: ";
    var sheetFound = false;
    var mainListRowFound = false;
    var quarterSheetRowFound = false;
    var formResponseFound = false;

    if (ui.alert("?האם למחוק", ui.ButtonSet.YES_NO) == ui.Button.YES) {

        rowToDelete = Number(ui.prompt("?מה מספר השורה אותה תרצי למחוק").getResponseText());

        if (ui.alert("????  " + quarterSheet.getRange("A" + rowToDelete).getDisplayValue() + " - " + quarterSheet.getRange("B" + rowToDelete).getDisplayValue() + "  ????", ui.ButtonSet.YES_NO) == ui.Button.YES) {

            name = quarterSheet.getRange("A" + rowToDelete).getDisplayValue();
            date = quarterSheet.getRange("B" + rowToDelete).getDisplayValue();

            //sheet
            if (hayavaSpreadsheet.getSheetByName(name + " - " + date)) {
                hayavaSpreadsheet.getSheetByName(name + " - " + date).copyTo(deletedSpreadsheet);
                deletedSpreadsheet.getSheetByName("Copy of " + name + " - " + date).setName(name + " - " + date);
                hayavaSpreadsheet.deleteSheet(hayavaSpreadsheet.getSheetByName(name + " - " + date));
                sheetFound = true;
            }

            //mainList
            for (var i = 2; i <= mainList.getLastRow(); i++) {
                if ((mainList.getRange("B" + i).getDisplayValue() == name || mainList.getRange("C" + i).getDisplayValue() == name) && mainList.getRange("G" + i).getDisplayValue() == date) {
                    mainList.deleteRow(i);
                    mainListRowFound = true;
                }
            }

            //quarterSheetRow
            for (var i = 2; i <= quarterSheet.getLastRow(); i++) {
                if (quarterSheet.getRange("A" + i).getDisplayValue() == name && quarterSheet.getRange("B" + i).getDisplayValue() == date) {
                    quarterSheet.deleteRow(i);
                    quarterSheetRowFound = true;
                }
            }

            //formResponse
            var formStartDate = "";
            for (var i = 0; i < responses.length; i++) {
                formStartDate = Utilities.formatDate(new Date(responses[i].getItemResponses()[5].getResponse()), Session.getScriptTimeZone(), "dd/MM/yyyy");
                if ((responses[i].getItemResponses()[0].getResponse() == name || responses[i].getItemResponses()[1].getResponse() == name) && formStartDate == date) {
                    FormApp.openById('1OBGW2vxpwAh915jRgdNrJEL77wPIcxc2DIhONCra0ew').deleteResponse(responses[i].getId());
                    formResponseFound = true;
                }
            }

            if (!sheetFound || !mainListRowFound || !quarterSheetRowFound || !formResponseFound) {
                if (!sheetFound) {
                    problemString += "  sheet  ";
                }

                if (!mainListRowFound) {
                    problemString += "  mainList  ";
                }

                if (!quarterSheetRowFound) {
                    problemString += "  quarterSheet  ";
                }

                if (!formResponseFound) {
                    problemString += "  FormResponse  ";
                }

                MailApp.sendEmail("omerbengal7+havayaNotifications@gmail.com", "problem with deleting function", problemString);
            } else {
                ui.alert("!המחיקה התבצעה בהצלחה");
            }

        }
    }

}






