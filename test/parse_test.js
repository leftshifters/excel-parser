'user strict';

var excelParser = require('../excelParser.js'),
    parser = {};

parser.parse_10000_xls = function(test) {
  test.expect(4);
  excelParser.parse({
    inFile: __dirname + '/files/custom.xls',
    worksheet: 1
  }, function(err, results) {
    var len = results.length,
        r1 = results[0][0],
        r10000 = results[results.length-1][0];

    test.strictEqual(len, 10000, 'number', 'Should be number');
    test.strictEqual(r1, 'Name', 'string', 'Should be string');
    test.strictEqual(r10000, 'Airhole', 'string', 'Should be string');
    test.ifError(err);
    test.done();
  });
};

parser.parse_10000_xlsx = function(test) {
  test.expect(4);
  excelParser.parse({
    inFile: __dirname + '/files/custom.xlsx',
    worksheet: 1
  }, function(err, results) {
    var len = results.length,
        r1 = results[0][0],
        r10000 = results[results.length-1][0];

    test.strictEqual(len, 10000, 'number', 'Should be number');
    test.strictEqual(r1, 'Name', 'string', 'Should be string');
    test.strictEqual(r10000, 'Airhole', 'string', 'Should be string');
    test.ifError(err);
    test.done();
  });
};

parser.searchRowsXls = function(test) {
  test.expect(0);
  excelParser.parse({
    inFile: __dirname + '/files/custom.xls',
    worksheet: 1,
    searchFor: {
      term: ["Airhole"],
      type: "strict"
    }
  }, function(err, results) {
    console.log(results.length);
    test.done();
  });
};

exports.parser = parser;