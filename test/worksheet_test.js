'user strict';

var excelParser = require('../excelParser.js');

exports.worksheets = {
	xlsWorksheets: function(test) {
		test.expect(4);
		excelParser.worksheets({inFile: __dirname + '/files/multi_worksheets.xls'}, function(err, worksheets){
			test.ok(worksheets, "Found all the worksheets");
			var s1 = worksheets[0].name;
			var i2 = worksheets[1].id;
			test.strictEqual(s1, 'Sheet1', 'string', "Should be string type.");
			test.strictEqual(i2, 2, 'number', "Should be number.");
			test.ifError(err);
			test.done();
		});
	},

	xlsxWorksheets: function(test) {
		test.expect(4);
		excelParser.worksheets({inFile: __dirname + '/files/multi_worksheets.xlsx'}, function(err, worksheets) {
			test.ok(worksheets, "Found all the worksheets");
			var s1 = worksheets[0].name;
			var i2 = worksheets[1].id;
			test.strictEqual(s1, 'Sheet1', 'string', "Should be string type.");
			test.strictEqual(i2, 2, 'number', "Should be number.");
			test.ifError(err);
			test.done();
		});
	}
};