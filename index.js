const fs = require('fs')
const xlsx = require('xlsx');

const excelWorkbook = xlsx.readFile('.\\users365.xlsx');

const arrayData = xlsx.utils.sheet_to_json(excelWorkbook.Sheets.users365);

const bibliofacileCharachtersNumber = {
  "Num": 3,
  "Nom": 61,
  "Prenom": 42,
  "mail": 521,
  "telephone": 29,
  "adresse": 81,
  "complement": 81,
  "ville": 81,
  "adresse secondaire": 81,
  "complement adresse secondaire": 81,
  "ville secondaire": 81,
  "groupe": 63
};

const propertiesMatchingFunctions = {
  "Num": ({acc}) => acc,
  "Nom": ({obj}) => obj["Nom"],
  "Prenom": ({obj}) => obj["Prénom"],
  "mail": ({obj}) => obj["Nom d’utilisateur principal"],
  "telephone": ({obj}) => "",
  "adresse": ({obj}) => "",
  "complement": ({obj}) => "",
  "ville": ({obj}) => "",
  "adresse secondaire": ({obj}) => "",
  "complement adresse secondaire": ({obj}) => "",
  "ville secondaire": ({obj}) => "",
  "groupe": ({obj}) => `${obj.Titre}${obj.Service}${obj.Bureau}`
};

const bibliofacileArray = arrayData.map((obj, acc) => {
  const formattedObject = {};
  Object.keys(propertiesMatchingFunctions).forEach((property, i) => {
    formattedObject[property] = propertiesMatchingFunctions[property]({obj, acc: acc+1})
  });
  return formattedObject;
});
