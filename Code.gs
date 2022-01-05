const CONFIG_SHEET_NAME = "Config";
const OUTPUT_LOCATION = "B5";
const TITLE_FILE = "B4";
const DOCUMENT_NAME_PREFIX = "WCK Resource List - ";
const SHEET = "Resources";
const HEADER = [
  "resource",
  "type_of_support",
  "include",
  "description",
  "covid_updates",
  "address",
  "website",
  "email",
  "phone",
  "hours",
  "age",
  "area",
  "fees",
  "referral",
  "services",
  "populations_served",
  "funding_options",
];

const style1 = {};
style1[DocumentApp.Attribute.BOLD] = true;

const h1style = {};
h1style[DocumentApp.Attribute.BOLD] = true;
h1style[DocumentApp.Attribute.UNDERLINE] = true;

const h2style = {};
h2style[DocumentApp.Attribute.BOLD] = false;
h2style[DocumentApp.Attribute.FONT_SIZE] = 18;

const no_bold = {};
no_bold[DocumentApp.Attribute.BOLD] = false;

const normal_text = {};
normal_text[DocumentApp.Attribute.FONT_FAMILY] = "Source Sans Pro";
normal_text[DocumentApp.Attribute.FONT_SIZE] = 10;

const highlighted = {};
highlighted[DocumentApp.Attribute.BACKGROUND_COLOR] = "#ffff00";
highlighted[DocumentApp.Attribute.BOLD] = true;

const footerstyle = {};
footerstyle[DocumentApp.Attribute.ITALIC] = true;
footerstyle[DocumentApp.Attribute.FONT_SIZE] = 8;

const after_paragraph_spacing = 10;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Resource Guide")
    .addItem("Create Resource Guide Doc", "main")
    .addToUi();
}

function parseRow(row) {
  let out = Object();
  for (let i = 0; i < HEADER.length; i++) {
    out[HEADER[i]] = row[i];
  }
  return out;
}

function parseRows(values) {
  let out = [];
  for (let row = 0; row < values.length; row++) {
    out.push(parseRow(values[row]));
  }
  return out;
}

function parseResourceTypes(parsed_rows) {
  let resource_types = [];
  let indices = [];
  for (let i = 0; i < parsed_rows.length; i++) {
    if (!resource_types.includes(parsed_rows[i]["type_of_support"])) {
      resource_types.push(parsed_rows[i]["type_of_support"]);
      indices.push(i);
    }
  }
  return indices;
}

function getSheetData() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  let values = sheet.getDataRange().getValues();
  let header = values[0];
  let body = values.slice(1);
  Logger.log(header);
  let rows = parseRows(body);
  let resource_type_indices = parseResourceTypes(rows);
  Logger.log(resource_type_indices);
  return { indices: resource_type_indices, body: rows };
}

function removeRows(info_table, row, removable_rows, first_row) {
  let removed = 0;
  for (let k = 0; k < removable_rows.length; k++) {
    if (!row[removable_rows[k]]) {
      info_table.removeRow(k + first_row - removed);
      removed++;
    }
  }
  return removed;
}

function getConfig() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  var out = {};
  out["title_file"] = sheet.getRange(TITLE_FILE).getValue();
  out["output_location"] = sheet.getRange(OUTPUT_LOCATION).getValue();
  return out;
}

function createDoc() {
  config = getConfig();
  // Create Doc
  doc_name = DOCUMENT_NAME_PREFIX + Date().toString();
  var doc = DocumentApp.create(doc_name);
  Logger.log("Created file named %s with id %s", doc.getName(), doc.getId());
  Logger.log("File url: %s", doc.getUrl());

  // Move to folder
  var folder = DriveApp.getFolderById(config.output_location);
  Logger.log("Located folder %s named %s", folder.getId(), folder.getName());
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  Logger.log("Moved file %s to folder %s", doc.getName(), folder.getName());

  // set doc settings
  let body = doc.getBody();
  let attributes = {};
  attributes[DocumentApp.Attribute.FONT_FAMILY] = "Source Sans Pro";

  // Add title page
  let title_image = DriveApp.getFileById(config.title_file).getBlob();
  body.appendImage(title_image).setWidth(600).setHeight(320);
  body.appendPageBreak();

  // Add TOC text
  body
    .appendParagraph("Contents")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setAttributes(h1style);
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.TITLE, attributes);
  body
    .appendParagraph(
      "[Instructions: insert table of contents and then add page numbers]"
    )
    .setAttributes(highlighted);
  body.appendPageBreak();

  // Get Data
  let data = getSheetData();

  // Add data
  for (let i = 0; i < data.body.length; i++) {
    let row = data.body[i];
    if (row["include"]) {
      // TYPE OF SUPPORT
      if (data.indices.includes(i)) {
        let support_heading = body.appendParagraph(row["type_of_support"]);
        support_heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        support_heading.setAttributes(h1style);
      }
      // END TYPE

      // RESOURCE NAME
      let resource_title = body.appendParagraph(row["resource"]);
      resource_title
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAttributes(h2style);
      // END RESOURCE NAME

      // CONTACT INFO & SERVICES
      info_table_cells = [["", "Contact Information"]];
      info_table_cells.push(["Phone", row.phone]);
      info_table_cells.push(["Website", row.website]);
      info_table_cells.push(["Email", row.email]);
      info_table_cells.push(["Address", row.address]);
      info_table_cells.push(["Hours", row.hours]);
      info_table_cells.push(["", "Service Information"]);
      info_table_cells.push(["Area", row.area]);
      info_table_cells.push(["Ages", row.age]);
      info_table_cells.push(["Populations served", row.populations_served]);
      info_table_cells.push(["Fees", row.fees]);
      info_table_cells.push(["Referral", row.referral]);
      info_table_cells.push(["Funding options", row.funding_options]);
      //       info table
      info_table = body.appendTable(info_table_cells);
      let table_width = 468;
      let col0_width = 108;
      let col1_width = table_width - col0_width;
      info_table.setColumnWidth(0, col0_width);
      info_table.setColumnWidth(1, col1_width);
      info_table.setAttributes(no_bold);
      info_table.getCell(0, 1).setAttributes(style1);
      info_table.getCell(2, 1).editAsText().setLinkUrl(row.website);
      info_table
        .getCell(3, 1)
        .editAsText()
        .setLinkUrl("mailto:" + row.email);
      info_table.getCell(6, 1).setAttributes(style1);

      // remove blank rows
      removed = removeRows(
        info_table,
        row,
        ["phone", "website", "email", "address", "hours"],
        1
      );
      removed = removeRows(
        info_table,
        row,
        [
          "area",
          "ages",
          "populations_served",
          "fees",
          "referral",
          "funding_options",
        ],
        7 - removed
      );

      // bold left column
      for (let j = 1; j < info_table.getNumRows(); j++) {
        info_table.getRow(j).getCell(0).setAttributes(style1);
      }
      // END CONTACT INFO

      // FULL NOTES
      body.appendParagraph("Description")
      .setAttributes(style1)
      .setSpacingBefore(after_paragraph_spacing);
      body
        .appendParagraph(row.description)
        .setHeading(DocumentApp.ParagraphHeading.NORMAL)
        .setSpacingAfter(after_paragraph_spacing);

      // SERVICES
      body.appendParagraph("Services").setAttributes(style1);
      body
        .appendParagraph(row["services"])
        .setHeading(DocumentApp.ParagraphHeading.NORMAL)
        .setSpacingAfter(after_paragraph_spacing);
      // END SERVICES

      // COVID UPDATES
      if (row.covid_updates) {
        body.appendParagraph("COVID-19 Updates").setAttributes(style1);
        body
          .appendParagraph(row.covid_updates)
          .setHeading(DocumentApp.ParagraphHeading.NORMAL);
      }
      body.appendPageBreak();

      // if (data.indices.includes(i+1)) {
      //   body.appendPageBreak()
      // }
    }
  }

  // FOOTER
  doc
    .addFooter()
    .appendParagraph(
      "WCK Family Resources List. Last updated " + Date().toString()
    )
    .setAttributes(footerstyle);

  // APPLY STYLES
  paragraphs = body.getParagraphs();
  for (j in paragraphs) {
    paragraphs[j].setAttributes(attributes);
    if (paragraphs[j].getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
      paragraphs[j].setAttributes(normal_text);
    }
  }
  Logger.log(doc.getUrl());

  body.setAttributes(attributes);
  return doc;
}

function showAlert(text) {
  alerthtml = HtmlService.createHtmlOutput(
    '<a href="' + text + '" target="_blank">See new doc</a>'
  )
    .setHeight(100)
    .setWidth(200);
  SpreadsheetApp.getUi().showModalDialog(alerthtml, "New doc created");
  return;
}

function main() {
  let doc = createDoc();
  showAlert(doc.getUrl());
}
