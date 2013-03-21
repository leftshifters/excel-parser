'user strict';

var excelParser = require('../excelParser.js');

exports.skipTest = {
	xlsSkipFalse: function(test) {
		test.expect(4);
		excelParser.parse({
			inFile: __dirname + '/files/empty_fields.xls',
			worksheet: 1,
			skipEmpty: false
		}, function(err, records) {
			test.ok(records, "Successfully parsed empty_fields.xls");
			var len = records[2].length;
			var skip = records[2][2];
			test.strictEqual(len, 5, 'number', "Found 5 records");
			test.strictEqual(skip, '', 'string', "Found empty");
			test.ifError(err);
			test.done();
		});
	},

	xlsSkipTrue: function(test) {
		test.expect(4);
		excelParser.parse({
			inFile: __dirname + '/files/empty_fields.xls',
			worksheet: 1,
			skipEmpty: true
		}, function(err, records) {
			test.ok(records, "Successfully parsed empty_fields.xls");
			var len = records[2].length;
			var skip = records[2][2];
			test.strictEqual(len, 4, 'number', "Found 4 records and skip one");
			test.strictEqual(skip, 'neque.vitae.semper@consectetuer.edu', 'string', "Found neque.vitae.semper@consectetuer.edu in place of empty");
			test.ifError(err);
			test.done();
		});
	},

	xlsxSkipFalse: function(test) {
		test.expect(4);
		excelParser.parse({
			inFile: __dirname + '/files/empty_fields.xlsx',
			worksheet: 1,
			skipEmpty: false
		}, function(err, records){
			test.ok(records, "Successfully parsed empty_fields.xls");
			var len = records[2].length;
			var skip = records[2][2];
			test.strictEqual(len, 5, 'number', "Found 5 records");
			test.strictEqual(skip, '', 'string', "Found empty");
			test.ifError(err);
			test.done();
		});
	},

	xlsxSkipTrue: function(test) {
		test.expect(4);
		excelParser.parse({
			inFile: __dirname + '/files/empty_fields.xlsx',
			worksheet: 1,
			skipEmpty: true
		}, function(err, records){
			test.ok(records, "Successfully parsed empty_fields.xls");
			var len = records[2].length;
			var skip = records[2][2];
			test.strictEqual(len, 4, 'number', "Found 4 records and skip one");
			test.strictEqual(skip, 'neque.vitae.semper@consectetuer.edu', 'string', "Found neque.vitae.semper@consectetuer.edu in place of empty");
			test.ifError(err);
			test.done();
		});
	}
};