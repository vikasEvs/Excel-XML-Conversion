// Import Library to use in the code
const fs = require("fs");
const xlsx = require("xlsx");

// Read the file with name test.xlsx
const workbook = xlsx.readFile("clientEpo.xlsx");

/** Select sheet where data is present, which we want to use */
const sheet = workbook.Sheets["Complete Data"];

/** Converting sheet data to JSON data */
const jsonData = xlsx.utils.sheet_to_json(sheet);

function parsePatentNumber(patentNumber) {
  const match = patentNumber.match(/^([A-Z]+)(\d+)([A-Z0-9]{2})$/);
  if (match) {
    return {
      auth: match[1],
      num: match[2],
      kind: match[3]
    };
  }
  return { auth: "", num: "", kind: "" };
}

function attributeValue(attribute) {
  return attribute === 'INVention' ? 'I' : attribute === 'ADDITional' ? 'A' : '';
}

function actionValue(action) {
  switch (action) {
    case "ADD":
    case "MODIFY ATTRIBUTE":
    case "CONFIRM UNCHANGED":
      return "A";
    case "DELETE":
      return "D";
    case "CIRCULATION":
      return "C"
    default:
      return "";
  }
}

function generateXML(jsonData) {
  /** Creating xmlString with initial text to come in XML file */
  let xmlString = `<patent-documents xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.epo.org/cpc/ecs/ox file://va99fp04/appresults$/ReClass/ox-v2.xsd" xmlns="http://www.epo.org/cpc/ecs/ox">`;

  let currentPatent = null;

  jsonData.forEach(entry => {
    if (entry['Doc#']) {
      if (currentPatent) {
        xmlString += '</allocations></document>';
      }
      const patentDetails = parsePatentNumber(entry['Patent Number']);
      xmlString += `<document auth="${patentDetails.auth}" num="${patentDetails.num}" kind="${patentDetails.kind}"><allocations>`;
      currentPatent = patentDetails;
    } else {
      const action = actionValue(entry.Action);
      const sourceAndGenOffice = 'source="H" gen-office="EP"';
      const posAndOrigin = action === 'D' ? '' : ' pos="L" origin="R"';
      xmlString += `<class symbol="${entry['Treated CPC']}" value="${attributeValue(entry.Attribute)}" ${sourceAndGenOffice}${posAndOrigin}>`
      xmlString += `<scheme scheme="CPC" />`
      xmlString += `<action value="${action}" />`
      xmlString += `</class>`;
    }
  });

  if (currentPatent) {
    xmlString += '</allocations></document>';
  }
  xmlString += "</patent-documents>";
  return xmlString;
}

let xmlData = generateXML(jsonData);

fs.writeFileSync("outputEpo.xml", xmlData);
