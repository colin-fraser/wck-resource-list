const OUTPUT_LOCATION = '1heJIjY0Dfo9ia5BFf9fj9AEeUHxnPGS8'
const DOCUMENT_NAME_PREFIX = 'WCK Resource List - '
const SHEET = 'Resources'
const HEADER = ['resource', 'type_of_support', 'include', 'notes', 'covid_updates', 'address',
    'website', 'email', 'phone', 'hours', 'age', 'area', 'fees', 'referral',
    'description', 'services', 'populations_served', 'funding_options']

var style1 = {}
style1[DocumentApp.Attribute.BOLD] = true

var footerstyle = {}
footerstyle[DocumentApp.Attribute.ITALIC] = true
footerstyle[DocumentApp.Attribute.FONT_SIZE] = 8

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Resource Guide')
        .addItem('Create Resource Guide Doc', 'main')
        .addToUi()
}

function parseRow(row) {
    let out = Object()
    for (let i = 0; i < HEADER.length; i++) {
        out[HEADER[i]] = row[i]
    }
    return out
}

function parseRows(values) {
    let out = []
    for (let row = 0; row < values.length; row++) {
        out.push(parseRow(values[row]))
    }
    return out
}

function parseResourceTypes(parsed_rows) {
    let resource_types = [];
    let indices = [];
    for (let i = 0; i < parsed_rows.length; i++) {
        if (!resource_types.includes(parsed_rows[i]["type_of_support"])) {
            resource_types.push(parsed_rows[i]["type_of_support"])
            indices.push(i)
        }
    }
    return indices
}

function getSheetData() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET)
    sheet.range
    let values = sheet.getDataRange().getValues();
    let header = values[0];
    let body = values.slice(1)
    Logger.log(header)
    let rows = parseRows(body);
    let resource_type_indices = parseResourceTypes(rows)
    Logger.log(resource_type_indices)
    return {indices: resource_type_indices, body: rows}
}


function createDoc() {
    // Create Doc
    doc_name = DOCUMENT_NAME_PREFIX + Date().toString()
    var doc = DocumentApp.create(doc_name)
    Logger.log('Created file named %s with id %s', doc.getName(), doc.getId())
    Logger.log('File url: %s', doc.getUrl())

    // Move to folder
    var folder = DriveApp.getFolderById(OUTPUT_LOCATION)
    Logger.log('Located folder %s named %s', folder.getId(), folder.getName())
    DriveApp.getFileById(doc.getId()).moveTo(folder)
    Logger.log('Moved file %s to folder %s', doc.getName(), folder.getName())

    // set doc settings
    let body = doc.getBody()
    let attributes = {};
    attributes[DocumentApp.Attribute.MARGIN_LEFT] = 72;
    attributes[DocumentApp.Attribute.MARGIN_RIGHT] = 72;
    attributes[DocumentApp.Attribute.INDENT_FIRST_LINE] = 0;
    attributes[DocumentApp.Attribute.INDENT_START] = 0;
    attributes[DocumentApp.Attribute.INDENT_END] = 0;
    attributes[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    body.setAttributes(attributes);

    // Get Data
    let data = getSheetData();

    // Add text
    body
        .appendParagraph("Resource List")
        .setHeading(DocumentApp.ParagraphHeading.TITLE)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("[Instructions: insert table of contents and then add page numbers]")
    body.appendPageBreak()


    // Add data
    for (let i = 0; i < data.body.length; i++) {

        let row = data.body[i];
        if (row['include']) {
            if (data.indices.includes(i)) {
                let support_heading = body.appendParagraph(row["type_of_support"])
                support_heading.setHeading(DocumentApp.ParagraphHeading.HEADING1)
            }

            let resource_title = body.appendParagraph(row['resource']);
            resource_title.setHeading(DocumentApp.ParagraphHeading.HEADING2)
            body
                .appendParagraph('Description')
                .setAttributes(style1)
            
            body
              .appendParagraph(row['description'])
              .setHeading(DocumentApp.ParagraphHeading.NORMAL)
            
            body
                .appendParagraph('Services')
                .setAttributes(style1)
            body
                .appendParagraph(row['services'])
                .setHeading(DocumentApp.ParagraphHeading.NORMAL)
            
            info_table_cells = [['', 'Contact Information']]
            info_table_cells.push(['Phone', row.phone])
            info_table_cells.push(['Email', row.email])
            info_table_cells.push(['Address', row.address])
            info_table_cells.push(['Hours', row.hours])
            info_table = body.appendTable(info_table_cells)
            info_table.setColumnWidth(0, 72)
            info_table.setColumnWidth(1, 72*3)
            info_table.getCell(0,1).setAttributes(style1)
            for (let j = 1; j < info_table.getNumRows(); j++) {
              info_table
                .getRow(j)
                .getCell(0)
                .setAttributes(style1)
            }
          
          if (data.indices.includes(i+1)) {
            body.appendPageBreak()
          }

        }
    }
    doc
      .addFooter()
      .appendParagraph("WCK Resource List. Last updated " + Date().toString())
      .setAttributes(footerstyle)
    Logger.log(doc.getUrl())
    return doc
}

function main() {
    let doc = createDoc()
    SpreadsheetApp
        .getUi()
        .alert(doc.getUrl())
}