// Import Library to use in the code
const fs = require("fs");
const xlsx = require("xlsx");

function parsePatentNumber(patentNumber) {
  // Extract the first two characters as the country code (auth)
  const auth = patentNumber.substring(0, 2);

  // Extract the remaining part after the country code
  const remaining = patentNumber.substring(2);

  // Use a regular expression to split the remaining part into num and kind
  const match = remaining.match(/^([A-Z]*\d+[A-Z]*)([A-Z]+\d*[A-Z]*)$/);

  if (match) {
    return {
      auth: auth,
      num: match[1],
      kind: match[2]
    };
  }

  return { auth: "", num: "", kind: "" };
}

function attributeValue(attribute) {
  attribute = attribute.toLowerCase();
  return attribute === "invention"
    ? "I"
    : attribute === "additional" ? "A" : "";
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
      return "C";
    default:
      return "";
  }
}

// Read the file with file name epoClient.xlsx
const workbook = xlsx.readFile("RP10453- Complete XML revised.xlsx");

/** Select sheet where data is present, which we want to use */
const sheet = workbook.Sheets["Complete Data"];
/** Converting sheet data to JSON data */
const jsonData = xlsx.utils.sheet_to_json(sheet);

function generateXML(jsonData) {
  /** Creating xmlString with initial text to come in XML file */
  let xmlString = `<patent-documents xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.epo.org/cpc/ecs/ox file://va99fp04/appresults$/ReClass/ox-v2.xsd" xmlns="http://www.epo.org/cpc/ecs/ox">`;

  let currentPatent = null;
  jsonData.forEach(entry => {
    if (entry["Doc#"]) {
      // console.log("entry[Doc#]-->>", entry["Doc#"]);
      if (currentPatent) {
        xmlString += "</allocations></document>";
      }
      // console.log("KIND CODE-->>",entry["Patent Number with kind code"]);
      const patentDetails = parsePatentNumber(
        entry["Patent Number with kind code"]
      );
      xmlString += `<document auth="${patentDetails.auth}" num="${patentDetails.num}" kind="${patentDetails.kind}"><allocations>`;
      currentPatent = patentDetails;
    } else {
      const action = actionValue(
        entry["Action - ADD/DELETE/CONFIRM UNCHANGED/MODIFY ATTRIBUTE"]
      );
      const sourceAndGenOffice = 'source="H" gen-office="EP"';
      const posAndOrigin = action === "D" ? "" : ' pos="L" origin="R"';
      let symbol =
        entry[
          "Treated CPC: Please delete incorrect CPC, confirm correct CPC, add missing CPC on next blank row beneath current classification and specifiy CPC action in column I"
        ];
      symbol = symbol ? symbol.split(" ").join("") : "";
      xmlString += `<class symbol="${symbol.trim()}" value="${attributeValue(
        entry["Attribute INVention, ADDITional as proposed output"]
      )}" ${sourceAndGenOffice}${posAndOrigin}>`;
      xmlString += `<scheme scheme="CPC" />`;
      xmlString += `<action value="${action}" />`;
      xmlString += `</class>`;
    }
  });

  if (currentPatent) {
    xmlString += "</allocations></document>";
  }
  xmlString += "</patent-documents>";
  return xmlString;
}

let xmlData = generateXML(jsonData);

fs.writeFileSync("RP10453- Complete XML revised.xml", xmlData);
