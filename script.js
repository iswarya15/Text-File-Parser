var xlsx = require('xlsx');

var workbook = xlsx.readFile('Stock.xlsx');

var worksheet = workbook.Sheets['Sheet1'];
// returns an array of json objects(each row in excel)
var data = xlsx.utils.sheet_to_json(worksheet);

console.log(worksheet)

console.log(data)

var newData = data.map((record) => {
   //Parse Stock Name
   const regex = /[\d]/g;
   var firstDigitIndex = record.Name.search(regex);
   record.StockName = record.Name.substring(0, firstDigitIndex - 1);

   //Parse Old price
   const rejexOpenBracket = /[(]/g;
   var firstBracketIndex = record.Name.search(rejexOpenBracket);
   record.OldPrice = record.Name.substring(firstDigitIndex, firstBracketIndex);

   //Parse Old price year
   const rejexCloseBracket = /[)]/g;
   var firstCloseBracketIndex = record.Name.search(rejexCloseBracket);
   record.OldYear = record.Name.substring(firstBracketIndex + 1, firstCloseBracketIndex);

   let removeUnusedString = record.Name.slice(firstCloseBracketIndex + 2);
   //Parse new price
   let secondBracketIndex = removeUnusedString.search(rejexOpenBracket);
   record.NewPrice = removeUnusedString.substring(0, secondBracketIndex);

   //Parse new year
   let secondCloseBracketIndex = removeUnusedString.search(rejexOpenBracket);
   record.newYear = removeUnusedString.substring(secondCloseBracketIndex + 1, removeUnusedString.length - 1);
   return record;


});



console.log(newData)
// create new excel workbook aka file
var newWorkBook = xlsx.utils.book_new();
var newWorkSheet = xlsx.utils.json_to_sheet(newData);

//append new sheet in created workbook
xlsx.utils.book_append_sheet(newWorkBook, newWorkSheet, "New Data");

xlsx.writeFile(newWorkBook, 'New Stock File.xlsx');


