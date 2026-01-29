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
const workbook = xlsx.readFile("RP11836-PB2 QC result for xml-Updated Kind codes 2.xlsx");
 
/** Select sheet where data is present, which we want to use */
const sheet = workbook.Sheets["Complete Data"];
/** Converting sheet data to JSON data */
const jsonData = xlsx.utils.sheet_to_json(sheet);
 
function generateXML(jsonData) {
  /** Creating xmlString with initial text to come in XML file */
  let xmlString = `<patent-documents xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.epo.org/cpc/ecs/ox file://va99fp04/appresults$/ReClass/ox-v2.xsd" xmlns="http://www.epo.org/cpc/ecs/ox">`;
 
  const COL_CPC_TREATED =
    'Treated CPC: Please delete incorrect CPC, confirm correct CPC, add missing CPC on next blank row beneath current classification and specifiy CPC action in column I';
 
  const COL_CSET =
    'Treated CPC C-set: Please delete incorrect CPC, confirm correct CPT, add missing CPC on next blank row beneath current classification and specifiy CPC action in column H';
 
  const COL_ACTION =
    'Action - ADD/DELETE/CONFIRM UNCHANGED/MODIFY ATTRIBUTE';
 
  const COL_VALUE =
    'Attribute INVention, ADDITional as proposed output';
 
  // Robust symbol cleaner: strips spaces, NBSP, zero-width chars
  const cleanSymbol = (s) =>
    (s == null ? "" : String(s))
      .replace(/[\u200B-\u200D\uFEFF]/g, "") // zero-width
      .replace(/\u00A0/g, " ")               // NBSP -> space
      .replace(/\s+/g, "")                   // all whitespace
      .trim();
 
  let currentPatent = null;
 
  jsonData.forEach(entry => {
    if (entry["Doc#"]) {
      // Start of a new patent document
      if (currentPatent) {
        xmlString += "</allocations></document>";
      }
      const patentDetails = parsePatentNumber(
        entry["Patent Number with kind code"]
      );
      xmlString += `<document auth="${patentDetails.auth}" num="${patentDetails.num}" kind="${patentDetails.kind}"><allocations>`;
      currentPatent = patentDetails;
 
    } else {
      // ===== CLASS logic (unchanged except: skip when empty) =====
      const action = actionValue(entry[COL_ACTION]);
      const sourceAndGenOffice = 'source="H" gen-office="EP"';
      const posAndOrigin = action === "D" ? "" : ' pos="L" origin="R"';
 
      const symbol = cleanSymbol(entry[COL_CPC_TREATED]);
 
      // Only emit a <class> if we actually have a symbol
      if (symbol) {
        xmlString += `<class symbol="${symbol}" value="${attributeValue(entry[COL_VALUE])}" ${sourceAndGenOffice}${posAndOrigin}>`;
        xmlString += `<scheme scheme="CPC" />`;
        xmlString += `<action value="${action}" />`;
        xmlString += `</class>`;
      }
 
      // ===== C-set logic (unchanged) =====
      let csetRaw = entry[COL_CSET];
 
      // Support multiple "cells" by allowing arrays OR newline-separated strings
      const csetCells = Array.isArray(csetRaw)
        ? csetRaw
        : (typeof csetRaw === "string" && csetRaw.trim().length > 0)
            ? csetRaw.split(/\r?\n/).filter(s => s.trim().length > 0)
            : [];
 
      if (csetCells.length > 0) {
        csetCells.forEach(cellStr => {
          // Split this cell into class symbols by comma, strip spaces etc.
          const classSymbols = String(cellStr)
            .split(",")
            .map(s => cleanSymbol(s))
            .filter(s => s.length > 0);
 
          if (classSymbols.length === 0) return;
 
          xmlString += `<cset>`;
 
          classSymbols.forEach((sym, idx) => {
            const rank = idx + 1;
            xmlString += `<rank rank="${rank}">`;
 
            if (rank === 1) {
              // Rank 1 with attributes + action
              xmlString += `<class symbol="${sym}" value="${attributeValue(entry[COL_VALUE])}" ${sourceAndGenOffice}${posAndOrigin}>`;
              xmlString += `<scheme scheme="CPC" />`;
              xmlString += `<action value="${action}" />`;
              xmlString += `</class>`;
            } else {
              // Rank 2+ minimal form
              xmlString += `<class symbol="${sym}">`;
              xmlString += `<scheme scheme="CPC" />`;
              xmlString += `</class>`;
            }
 
            xmlString += `</rank>`;
          });
 
          xmlString += `</cset>`;
        });
      }
      // ===== END C-set logic =====
    }
  });
 
  if (currentPatent) {
    xmlString += "</allocations></document>";
  }
  xmlString += "</patent-documents>";
 
  return xmlString;
}
let xmlData = generateXML(jsonData);
 
fs.writeFileSync("RP11836-PB2 QC result for xml-Updated Kind codes 2.xml", xmlData);
