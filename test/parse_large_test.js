'user strict';

var excelParser = require('../excelParser.js');

exports.parseLargeXls = function(test) {
	test.expect(3);
	excelParser.parse({
		inFile: __dirname + '/files/large.xls',
		worksheet: 1
	}, function(err, records) {
		test.ok(records, "Successfully parsed all worksheets from multi_worksheets.xls");
		var a = records.length;
		test.strictEqual(a, 201, 'number', "Found 201 records");
		test.ifError(err);
		test.done();
	});
};

exports.parseLargeXlsx = function(test) {
  test.expect(3);
  excelParser.parse({
    inFile: __dirname + '/files/large.xlsx',
    worksheet: 1
  }, function(err, records) {
    test.ok(records, "Successfully parsed all worksheets from multi_worksheets.xls");
    var a = records.length;
    test.strictEqual(a, 201, 'number', "Found 201 records");
    test.ifError(err);
    test.done();
  });
};