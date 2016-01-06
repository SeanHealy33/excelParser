var xlsx = require('node-xlsx');
var fs = require("fs");
var path = require('path');

var args = process.argv.slice(2);
var obj = xlsx.parse(args[0]); // parses a file
var excelDir = '../excelFiles';
console.log(__dirname+'..');

fs.writeFileSync(path.join(__dirname, "../excel.json"), JSON.stringify(obj));

console.log("Building file: This will take a while");
for(var i = 0; i < obj.length; i++ ){
	process.stdout.write(String(i/obj.length * 100).substring(0, 4) + "% Complete" + "\r");
	var buffer = xlsx.build([{name: (obj[i].name).toLowerCase(), data: obj[i].data}]);
	fs.writeFileSync(path.join(__dirname, excelDir + "/" + (obj[i].name).toLowerCase() + ".xlsx"), buffer);
}
