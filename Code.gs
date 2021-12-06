const OUTPUT_LOCATION = '1heJIjY0Dfo9ia5BFf9fj9AEeUHxnPGS8'
const DOCUMENT_NAME_PREFIX = 'WCK Resource List - '
const SHEET = 'Resources'
const HEADER = ['resource', 'type_of_support', 'include', 'phone', 'notes', 'covid_updates', 'address',
    'website', 'email', 'phone', 'hours', 'age', 'area', 'fees', 'referral',
    'description', 'services', 'populations_served', 'funding_options']

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
    let header = body.appendParagraph("Resource List")
    header.setHeading(DocumentApp.ParagraphHeading.TITLE);
    Logger.log('File url: %s', doc.getUrl())

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
            body.appendParagraph('Description').setHeading(DocumentApp.ParagraphHeading.HEADING3)
            let description = body.appendParagraph(row['description']);
            description.setHeading(DocumentApp.ParagraphHeading.NORMAL)
            body
                .appendParagraph('Services')
                .setHeading(DocumentApp.ParagraphHeading.HEADING3)
            doc.getOut
            body
                .appendParagraph(row['services'])
                .setHeading(DocumentApp.ParagraphHeading.NORMAL)

        }
    }
    return doc
}

function main() {
    let doc = createDoc()
    SpreadsheetApp
        .getUi()
        .alert(doc.getUrl())
}