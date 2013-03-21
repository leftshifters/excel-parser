'user strict';

var excelParser = require('../excelParser.js');

exports.searchLoose = function(test) {
	test.expect(4);
	excelParser.parse({
		inFile: __dirname + '/files/multi_worksheets.xls',
		worksheet: 1,
		searchFor: {
			term: [".edu", "lani"],
			type: "loose"
		}
	}, function(err, records) {
		test.ok(records, "Successfully parsed all worksheets from multi_worksheets.xls");
		var a = records.length;
		var e = records[0][0];
		test.strictEqual(a, 3, 'number', "Found 3 records");
		test.strictEqual(e, 'Lani', 'string', 'Found name "Lani" or "lani"');
		test.ifError(err);
		test.done();
	});
};