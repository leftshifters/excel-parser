var _  = require('underscore'),
    exec = require('child_process').exec,
    fs = require('fs'),
    path = require('path'),
    temp = require('temp');

//
// Excel parser
//
var excelParser = {
  worksheets: function(options, callback) {
    var skip=-1, cmd, worksheets, xlsFlag=false, response=[];
    if(
      !options || typeof(options) !== 'object' || typeof(options) === 'function'
      || !callback || typeof(callback) !== 'function'
      ) {
      throw new Error('"parse" required two arguments. parse(options, callback)');
    }
    if(!options.inFile) {
      return callback(new Error('File not found'));
    }
    _fileType(options.inFile, function(err, type) {
      var args = ['-x', options.inFile, '-W'];
      if(err) {
        callback(err);
      } else {
        if(type === 'xls') {
          cmd = "xls2csv ";
          xlsFlag = true;
          args.push('-q', '-f');
        } else if(type === 'xlsx') {
          cmd = "python " + __dirname + "/xlsx2csv.py ";
        } else {
          callback(new Error("Supports only xls and xlsx formats"));
        }

        _execute(cmd, args, function(err, stdout) {
          if(err) {
            callback(err);
          } else {
            worksheets = _.compact(stdout.split(/\n/));
            if(worksheets) {
              if(xlsFlag) {
                worksheets = worksheets.splice(1);
                worksheets = _.map(worksheets,function(sheet, index){return{'id': index+1, 'name': sheet}});
                callback(null, worksheets);
              } else {
                callback(null, JSON.parse(worksheets));
              }
            } else {
              callback(new Error("Not found any worksheet in given document"));
            }
          }
        });
      }
    });
  },
  parse: function(options, callback) {
    var self=this, cmd, worksheets, holder =[], i=0, response={};
    if(
      !options || typeof(options) !== 'object' || typeof(options) === 'function'
      || !callback || typeof(callback) !== 'function'
      ) {
      throw new Error('"parse" accepts two arguments. parse(options, callback)');
    }
    if(!options.inFile) {
      return callback(new Error('File not found'));
    }
    _fileType(options.inFile, function(err, type) {
      var args = ['-x', options.inFile];
      if(err) {
        callback(err);
      } else if(type === "xls") {
        cmd = "xls2csv ";
        args.push('-q', '-f');
      } else if(type === "xlsx") {
        cmd = "python " + __dirname + "/xlsx2csv.py ";
      } else {
        return callback(new Error("Supports only xls and xlsx formats"));
      }
      if(!options.worksheet){
        self.worksheets({inFile: options.inFile}, function(err, worksheets) {
          if(err){
            callback(err);
          } else {
            var _customLoop = function(worksheets){
              if(_.indexOf(args, '-n') < 0) {
                args.push('-n', worksheets[self._index].id);
              } else {
                args.splice(_.indexOf(args, '-n') + 1, 1, worksheets[self._index].id);
              }
              _pickRecords(cmd, args, options, function(err, records) {
                if(err) {
                  return callback(err);
                } else {
                  if(!_.isEmpty(records)) {
                    holder.push(records);
                  }
                  self._index++;
                  if(self._index < worksheets.length) {
                    _customLoop(worksheets);
                  } else {
                    callback(null, holder);
                  }
                }
              });
            };
            _customLoop(worksheets);
          }
        });
      } else {
        if(typeof(options.worksheet) === 'number') {
          args.push('-n', options.worksheet);
        } else if(typeof(options.worksheet) === 'string') {
          args.push('-w', '"'+ options.worksheet.trim() + '"');
        } else {
          return callback(new Error('"worksheet" parameter should be string or intiger'));
        }
        _pickRecords(cmd, args, options, function(err, records){
          if(err) {
            callback(err);
          } else {
            callback(null, records);
          }
        });
      }
    });
  },
  _index: 0,
  _searchIndex: 0
};

//
// Private methods
//
var _execute = function(cmd, args, callback) {
  exec(
    cmd + args.join().replace(/,/g, ' '),
    {cwd: __dirname},
    function(err, stdout, stderr) {
      if(err) return callback(new Error(err));
      if (stderr) return callback(new Error(stderr));
      callback(null, stdout);
  });
};

var _fileType = function(inFile, callback) {
  fs.exists(inFile, function(exists) {
    if(!exists) {
      callback(new Error('File does not exist'));
    } else {
      var type = path.extname(inFile).toLowerCase().slice(1);
      if(type !== 'xls' && type !== 'xlsx') {
        callback(new Error('Invalid file format'));
      } else {
        callback(null, type);
      }
    }
  });
};

var _csvParser = function(csvData, options, callback) {
  var results = [[]],
      records = [],
      match = null,
      skipEmpty = options.skipEmpty || false,
      searchFor = (
        options.searchFor &&
        typeof(options.searchFor === 'object') &&
        typeof(options.searchFor.term === 'object')
        ) ? options.searchFor.term : null,
      searchType = (
        options.searchFor &&
        typeof(options.searchFor.type === 'string') &&
        options.searchFor.type &&
        (
          options.searchFor.type.trim() === 'strict' ||
          options.searchFor.type.trim() === 'loose'
        )) ? options.searchFor.type.trim() : 'loose',
      searchPattern = null,
      skipEmpty = typeof(options.skipEmpty) === 'boolean' ? options.skipEmpty : false;

  var pattern = new RegExp((
    "(\\,|\\r?\\n|\\r|^)" +
    "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
    "([^\"\\,\\r\\n]*))"
  ), "gi");

  if(searchFor && searchType === 'strict') {
    searchPattern = _.map(searchFor, function(s) {
      return "\\b" + s + "\\b";
    });
    searchPattern = new RegExp('(' + searchPattern.join('|') + ')', "g");
  } else if(searchFor && searchType === 'loose') {
    searchPattern = new RegExp('(' + searchFor.join('|') + ')', "gi");
  }

  while(match = pattern.exec(csvData)) {
    var stringMatched = match[1];
    if(stringMatched.length && (stringMatched != ",")) {
      results.push([]);
    }
    if (match[2]) {
      var stringMatched = match[2].replace( new RegExp( "\"\"", "g" ), "\"");
    } else {
      var stringMatched = match[3];
    }

    if(skipEmpty) {
      if(!_.isEmpty(stringMatched)) {
        results[results.length-1].push(stringMatched);
      }
    } else {
      results[results.length-1].push(stringMatched);
    }
  }

  if(searchPattern) {
    var _customSearchLoop = function(searchIn) {
      if(!_.isEmpty(searchIn[excelParser._searchIndex])) {
        var arrString = searchIn[excelParser._searchIndex].join(' ');
        if(searchPattern.test(arrString)) {
          records.push(searchIn[excelParser._searchIndex]);
        }
      }
      excelParser._searchIndex++;
      if(excelParser._searchIndex < searchIn.length) {
        _customSearchLoop(searchIn);
      } else {
        callback(null, records);
      }
    };
    _customSearchLoop(results);
  } else {
    callback(null, results);
  }
};

var _pickRecords = function(cmd, args, options, callback) {
  var holder = [], i = 0;
  temp.mkdir('temp', function(err, dirPath) {
    if(err) {
      callback(err);
    } else {
      var csvFile = path.join(dirPath, 'convert.csv');
      args.push('-c', csvFile);
      _execute(cmd, args, function(err, stdout) {
        if(err) {
          return callback(err);
        } else {
          fs.readFile(csvFile, 'utf-8', function(err, data) {
            if(err) {
              callback(err);
            } else {
              _csvParser(data, options, function(err, records) {
                if(err) {
                  return callback(err);
                } else {
                  return callback(null, records);
                }
              });
            }
          });
        }
      });
    }
  });
};

module.exports = excelParser;