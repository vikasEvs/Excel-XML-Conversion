//import Library to use in the code
const fs = require("fs");
const xlsx = require("xlsx");

//read the file with name clientSadiq
const workbook = xlsx.readFile("SAD-PB3-24-Final 19 Nov.xlsx");

/** Select sheet where data is present, which we want to use */
const sheet = workbook.Sheets["All Data"];
// const sheet = workbook.Sheets[workbook.SheetNames[2]];

/** Converting sheet data to json data */
const data = xlsx.utils.sheet_to_json(sheet);

/** Creating xmlString with initial text to come in xml file */
let xmlString = `<exchange date="${new Date().toISOString()}" type="SADIQ" schema-version="0.4" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="sadiq-v1.3.xsd">`;

//Some global variable
let prevEdit = false;
let prevPro = false;
let prevDoc = false;
let prevApp = false;
let prevInf = false;

/** Object of Symbols which need to be converted into there html entity */
const symbolToEntity = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  "'": "&apos;",
  '"': "&quot;"
};

/** This function will replace the symbols with there entities */
function replaceSymbolsWithEntities(inputString) {
  return inputString.replace(
    /[&<>'"]/g,
    match => symbolToEntity[match] || match
  );
}

/** Here we will iterate to every row of excel to get data and insert into xml string */
for (let i = 0; i < data.length; i++) {
  /** First we are getting the data of Column "Fields" and "Components/data" */
  let fieldName = data[i]["Fields"]?data[i]["Fields"].trim():data[i]["Fields"];
  let fieldData = data[i]["Components/data"]?data[i]["Components/data"].trim():data[i]["Components/data"];

  /** We will append data to xml string from different column only if data will present in this column "Components/data" of the row*/
  if (fieldData) {
    fieldData = replaceSymbolsWithEntities(fieldData); // checking all the text if there will be any symbol it will convert to the entity and return the string
    if (i === 0 && fieldName === "PN") {
      prevDoc = true;
      xmlString += `<document pn="${fieldData}">`;
    } else if (fieldName === "PN") {
      if(prevApp){
        xmlString += `</app>`;
        prevApp = false;
      }
      if(prevEdit){
        xmlString += `</edit>`;
        prevEdit = false;
      }
      if(prevInf){
        xmlString += `</inf>`;
        prevInf = false;
      }
      if (prevDoc) {
        xmlString += "</document>";
        prevDoc = false;
      }
      if(!prevDoc){
        prevDoc = true;
        xmlString += `<document pn="${fieldData}">`;
      }
    }

    if (fieldName === "APP") {
      prevApp = true;
      xmlString += `<app>${fieldData}`;
    } else if (!fieldName && prevApp) {
      xmlString += `${fieldData}`;
    } else if (fieldName && prevApp) {
      xmlString += `</app>`;
      prevApp = false;
    }
    
    if (fieldName === "INF") {
      prevInf = true;
      xmlString += `<inf>${fieldData}`;
    } else if (!fieldName && prevInf) {
      xmlString += `${fieldData}`;
    } else if (fieldName && prevInf) {
      xmlString += `</inf>`;
      prevInf = false;
    }

    if (fieldName === "EDIT") {
      prevEdit = true;
      xmlString += "<edit>";
      const itemName = fieldData;
      const startValue = data[i]["Lower value"];
      const endValue = data[i]["Higher value"];
      xmlString += `<item name="${itemName}" start="${startValue}" end="${endValue}"/>`;
    } else if (!fieldName && prevEdit) {
      prevEdit = true;
      const itemName = fieldData;
      const startValue = data[i]["Lower value"];
      const endValue = data[i]["Higher value"];
      xmlString += `<item name="${itemName}" start="${startValue}" end="${endValue}"/>`;
    } else if (fieldName && fieldName !== "PN" && fieldName !== "APP" && fieldName !== "INF") {
      if (prevEdit) {
        xmlString += "</edit>";
        prevEdit = false;
      }
      if (
        fieldName === "PRO" &&
        data.length !== i + 1 &&
        !data[i + 1]["Fields"]
      ) {
        prevPro = true;
        xmlString += `<${fieldName.toLowerCase()}>${fieldData}`;
      } else {
        xmlString += `<${fieldName.toLowerCase()}>${fieldData}</${fieldName.toLowerCase()}>`;
      }
    } else if (!fieldName && prevPro) {
      if (data.length !== i + 1 && !data[i + 1]["Fields"]) {
        prevPro = true;
        xmlString += `${fieldData}`;
      } else {
        prevPro = false;
        xmlString += `${fieldData}</pro>`;
      }
    }
  } 
}

/** In last appending the closing tag in XML string */
if (prevApp) {
  xmlString += `</app>`;
}
if (prevEdit) {
  xmlString += `</edit>`;
}
if (prevInf) {
  xmlString += `</inf>`;
}
if (prevPro) {
  xmlString += `</pro>`;
}
if (prevDoc) {
  xmlString += "</document>";
}
xmlString += "</exchange>";

fs.writeFileSync("SAD-PB3-24-Final 19 Nov.xml", xmlString);
