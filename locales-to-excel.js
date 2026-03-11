#!/usr/bin/env node

import { readFileSync } from "fs";

import ExcelJS from "exceljs";
import { existsSync, readdirSync } from "node:fs";
import path from "path";

if (!existsSync("content")) {
  throw new Error(`Missing content folder`);
}

function flattenObject(obj, parentKey = "", result = {}) {
  for (let key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      const newKey = parentKey ? `${parentKey}.${key}` : key;
      const value = obj[key];

      if (typeof value === "object" && value !== null) {
        if (Array.isArray(value)) {
          value.forEach((item, index) => {
            const arrayKey = `${newKey}[${index}]`;

            if (typeof item === "object" && item !== null) {
              flattenObject(item, arrayKey, result);
            } else {
              result[arrayKey] = item; // keep primitive values as-is
            }
          });
        } else {
          flattenObject(value, newKey, result);
        }
      } else {
        result[newKey] = value;
      }
    }
  }
  return result;
}

const dictionary = {};
const languages = new Set();

readdirSync("./content", { withFileTypes: true, recursive: true })
  .filter((file) => file.isFile() && file.name.endsWith(".json"))
  .forEach((file) => {
    const languageCode = file.parentPath
      .replace("content/", "")
      .replace("content\\", "");

    const translation = readFileSync(path.join(file.parentPath, file.name), "utf8");


    const json = JSON.parse(translation);

    if (!dictionary[file.name]) {
      dictionary[file.name] = {};

    }
    dictionary[file.name][languageCode] = flattenObject(json);
    languages.add(languageCode);
  });



// Create a new workbook
const workbook = new ExcelJS.Workbook();

Object.keys(dictionary).forEach(file => {
  const worksheet = workbook.addWorksheet(file);

  const keys = Object.keys(dictionary[file]["en"]);

  const data = keys.map((key) => {
    const keyData = {
      keyName: key,
    };

    languages.forEach((l) => {
      keyData[l] = dictionary[file][l][key]?.toString() ?? "";
    });

    return keyData;
  });


  // Define columns
  worksheet.columns = [
    { header: "Key", key: "keyName", width: 90 },
    ...Array.from(languages).map((l) => ({ header: l, key: l, width: 50 })),
  ];

  // Add rows
  data.forEach((item) => {
    worksheet.addRow(item);
  });
});

// Write the workbook to a file
await workbook.xlsx.writeFile("translations.xlsx");
console.log("Excel file created!");
