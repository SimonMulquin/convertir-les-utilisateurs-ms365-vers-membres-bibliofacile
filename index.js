const fs = require('fs')
const xlsx = require('xlsx');
const { convertArrayToCSV } = require('convert-array-to-csv');

const excelWorkbook = xlsx.readFile('.\\users365.xlsx');

const arrayData = xlsx.utils.sheet_to_json(excelWorkbook.Sheets.users365);

const bibliofacileCharachtersNumber = {
  "Num": 3,
  "Nom": 61,
  "Prenom": 41,
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
const formatToCharachtersAmount = (str, property) => {
  if property == "Num" {


  }
  let string = "";
  if (typeof(str) == "string") {
    string = str;
  }
  const emptySpacesNumber = bibliofacileCharachtersNumber[property] - string.length - 1;
  const emptySpaces = (es="", acc=0) => {
    if (acc >= emptySpacesNumber) {
      return es;
    } else {
      return emptySpaces(`${es} `, acc+1);
    }
  }
  return ` ${string}${emptySpaces()}`; //one empty space as first character
};

const propertiesMatchingFunctions = {
  "Num": ({acc}) => formatToCharachtersAmount(acc, "Num"),
  "Nom": ({obj}) => formatToCharachtersAmount(obj["Nom"], "Nom"),
  "Prenom": ({obj}) => formatToCharachtersAmount(obj["Prénom"], "Prenom"),
  "mail": ({obj}) => formatToCharachtersAmount(obj["Nom d’utilisateur principal"], "mail"),
  "telephone": ({obj}) => formatToCharachtersAmount("", "telephone"),
  "adresse": ({obj}) => formatToCharachtersAmount("", "adresse"),
  "complement": ({obj}) => formatToCharachtersAmount("", "complement"),
  "ville": ({obj}) => formatToCharachtersAmount("", "ville"),
  "adresse secondaire": ({obj}) => formatToCharachtersAmount("", "adresse secondaire"),
  "complement adresse secondaire": ({obj}) => formatToCharachtersAmount("", "complement adresse secondaire"),
  "ville secondaire": ({obj}) => formatToCharachtersAmount("", "ville secondaire"),
  "groupe": ({obj}) => formatToCharachtersAmount(`${obj.Titre}${obj.Service}${obj.Bureau}`, "groupe")
};

const bibliofacileArray = arrayData.map((obj, acc) => {
  const formattedObject = {};
  Object.keys(propertiesMatchingFunctions).forEach((property, i) => {
    formattedObject[property] = propertiesMatchingFunctions[property]({obj, acc: acc+1})
  });
  return formattedObject;
});

const dataToWrite = convertArrayToCSV(bibliofacileArray)

fs.writeFile('membres.csv', dataToWrite, 'utf8', function (err) {
  if (err) {
    console.log('Some error occured - file either not saved or corrupted file saved.');
  } else{
    console.log('Fichier membres.csv généré !');
  }
});
