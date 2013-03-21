'user strict';

var excelParser = require('../excelParser.js');

exports.parseTest = {
	xlsParse: function(test) {
		test.expect(6);
		excelParser.parse({
			inFile: __dirname + '/files/multi_worksheets.xls',
			worksheet: 1
		}, function(err, records) {
			test.ok(records, "Successfully parsed all worksheets from multi_worksheets.xls");
			var a = records[0][0];
			var e = records[1][2];
			var i = records[2][3];
			var o = records[3][4];
			test.strictEqual(a, 'Names', 'string', "Found header Names");
			test.strictEqual(e, '96791', 'string', 'Found number 96791');
			test.strictEqual(i, 'neque.vitae.semper@consectetuer.edu', 'string', "Found email neque.vitae.semper@consectetuer.edu");
			test.strictEqual(o, '30-May-13', 'string', "Found date 30-May-13");
			test.ifError(err);
			test.done();
		});
	},

	xlsxParse: function(test) {
		test.expect(6);
		excelParser.parse({
			inFile: __dirname + '/files/multi_worksheets.xlsx',
			worksheet: 1
		}, function(err, records) {
			test.ok(records, "Successfully parsed all worksheets from multi_worksheets.xlsx");
			var a = records[0][0];
			var e = records[1][2];
			var i = records[2][3];
			var o = records[3][4];
			test.strictEqual(a, 'Names', 'string', "Found header Names");
			test.strictEqual(e, '96791', 'string', 'Found number 96791');
			test.strictEqual(i, 'neque.vitae.semper@consectetuer.edu', 'string', "Found email neque.vitae.semper@consectetuer.edu");
			test.strictEqual(o, '30-May-13', 'string', "Found date 30-May-13");
			test.ifError(err);
			test.done();
		});
	}
};