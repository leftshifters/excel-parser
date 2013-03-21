'user strict';

var excelParser = require('../excelParser.js');

exports.parseLarge = function(test) {
	test.expect(3);
	excelParser.parse({
		inFile: __dirname + '/files/large.xls',
		worksheet: 1
	}, function(err, records) {
		test.ok(records, "Successfully parsed all worksheets from multi_worksheets.xls");
		var a = records.length;
		test.strictEqual(a, 203, 'number', "Found 203 records");
		test.ifError(err);
		test.done();
	});
};