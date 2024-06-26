//import Library to use in the code
import fs from "fs";
import xlsx from 'xlsx';

//read the file with name clientAlloy
const workbook = xlsx.readFile("clientAlloy.xlsx");

/** Select sheet where data is present, which we want to use */
const sheet = workbook.Sheets["All Data"];
// const sheet = workbook.Sheets[workbook.SheetNames[2]];

/** Converting sheet data to json data */
const data = xlsx.utils.sheet_to_json(sheet);

/** Creating xmlString with initial text to come in xml file */
let xmlString = `<exchange date="${new Date().toISOString()}" type="ALLOYS" schema-version="1.2" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="alloys-v1.2.xsd">`;

//Some global variable
let prevEdit = false;
let prevDoc = false;
let prevKw = false;

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
  console.log("*******",data[i]);
  /** First we are getting the data of Column "Fields" and "Components/data" */
  let fieldName = data[i]["Fields"]? data[i]["Fields"].trim(): data[i]["Fields"];
  let fieldData = data[i]["Components/data"];

  /** We will append data to xml string from different column only if data will present in this column "Components/data" of the row*/
  if (fieldData) {
    console.log("FIELDNAME-->>",fieldName, fieldData);
    fieldData = replaceSymbolsWithEntities(fieldData); // checking all the text if there will be any symbol it will convert to the entity and return the string
    if (i === 0 && fieldName === "PN") {
      prevDoc = true;
      xmlString += `<document pn="${fieldData}">`;
    } else if (!prevDoc && fieldName === "PN") {
      prevDoc = true;
      xmlString += "</document>";
      xmlString += `<document pn="${fieldData}">`;
    }
    if(fieldName === "KW"){
      fieldData = fieldData.replace(/#/g, "&lt;br/&gt;");
    }
    if(fieldName === "TXT"){
      fieldData = fieldData.replace(/§\s*/g, '&lt;p&gt;').replace(/\n/g, '&lt;/p&gt;');
      fieldData += '&lt;/p&gt;'
    }
    if (fieldName === "EDIT" && fieldName !== "PRES" && fieldName !== "OPT") {
      prevEdit = true;
      xmlString += "<edit>";
      const itemName = fieldData;
      const startValue = data[i]["Min"];
      const endValue = data[i]["Max"];
      const isPresent = data[i]["Present"];
      xmlString += `<item name="${itemName}" start="${startValue}" end="${endValue}" present="${isPresent}"/>`;
    } else if (
      !fieldName &&
      prevEdit &&
      fieldName !== "PRES" &&
      fieldName !== "OPT"
    ) {
      console.log("1");
      prevEdit = true;
      const itemName = fieldData;
      const startValue = data[i]["Min"];
      const endValue = data[i]["Max"];
      const isPresent = data[i]["Present"];
      xmlString += `<item name="${itemName}" start="${startValue}" end="${endValue}" present="${isPresent}"/>`;
    } else if (
      fieldName &&
      fieldName !== "PN" &&
      fieldName !== "PRES" &&
      fieldName !== "OPT"
      ) {
      console.log("2");
      if (prevEdit) {
        xmlString += "</edit>";
        prevEdit = false;
      }
      xmlString += `<${fieldName.toLowerCase()}>${fieldData}</${fieldName.toLowerCase()}>`;
      if (data.length !== i + 1 && data[i + 1]["Fields"] === "PN") {
        prevDoc = false;
      }
    }
  } else if (data.length !== i + 1 && data[i + 1]["Fields"] === "PN") {
    console.log("3");
    prevDoc = false;
  }
}

/** In last appending the closing tag in xml string */
xmlString += "</document>";
xmlString += "</exchange>";

fs.writeFileSync("outputAlloy.xml", xmlString);
