// ignore_for_file: prefer_typing_uninitialized_variables
// ignore_for_file: depend_on_referenced_packages

import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

/// This contains low-level functions that directly interact with the Excel file.
class LowLevel {
  // Hold the Excel file currently being edited.
  late Excel _file;

  /// Create a new Excel file.
  LowLevel.createNew() {
    _file = Excel.createExcel();
  }

  /// Opening an existing Excel file.
  LowLevel.open({required String fileName}) {
    try {
      _file = Excel.decodeBytes(File(fileName).readAsBytesSync());
      print('File Open : $fileName');
    } catch (e) {
      throw Exception(
          "'$fileName' \nThis location could not be found. Please double-check the file name and file path before trying again.");
    }
  }

  /// Getting the maximum number of rows in the opened excel sheet.
  int get maxRow => _file[sheetName].maxRows;

  /// Getting the maximum number of columns in the opened excel sheet.
  int get maxCol => _file[sheetName].maxCols;

  /// Get the name of the opened excel sheet.
  String get sheetName => _file.getDefaultSheet()!;

  /// Setting the new user-defined name in the open Excel sheet.
  bool setSheet(String sheet) => _file.setDefaultSheet(sheet);

  /// Entering user defined data into a selected cell.
  void setCell({required String index, required String value}) {
    _file[sheetName].cell(CellIndex.indexByString(index)).value = value;
  }

  /// Get the data in the selected cell.
  String? getCell({required String index}) {
    return _file[sheetName].cell(CellIndex.indexByString(index)).value?.toString();
  }

  /// Save the opened Excel file.
  void saveFile({path = 'Output', name = 'final'}) {
    try {
      File(join(path + '/' + name + '.xlsx'))
        ..createSync(recursive: true)
        ..writeAsBytesSync(_file.save()!);
      print(
          'Total Column : ${_file[sheetName].maxCols}\nTotal Row : ${_file[sheetName].maxRows}\nFile Save : $path/$name.xlsx');
    } catch (e) {
      throw Exception(
          'The $name.xlsx file located on the $path/ path is currently being controlled by another software.'
          ' Close the $name.xlsx file before trying again. Or change the path or file name you want to save the file to.');
    }
  }
}

/// Excel files combination and data processing.
class Processing {
  late LowLevel _file;

  Processing() {
    _file = LowLevel.createNew(); // Create a new Excel file.
  }

  /// Reorganizing the data into the specified order.
  ///
  /// files: List of directories containing files to be edited.
  ///
  /// titleBar: The title order of the titles in the title bar.
  ///
  /// sequence: The order in which the data should be entered in each column.
  ///
  /// prefixes: Set this to true to run the prefix function if you want to add characters after a numeric value or change the characters after a numeric value. Data should be entered for the parameters below only if this is true.
  ///
  /// column: The name of the column in which to add characters after a numeric value or change characters after a numeric value.
  ///
  /// swap: If there are no letters after a numeric value, the letter or word to add.
  ///
  /// notaion: The letters to change after a numeric value and the word to replace them with. for Example ['k', ' Kohms', 'M', ' Mohms'];
  void sheetsConfig({
    required List<String> files,
    required List<String> titleBar,
    required Map<String, Set<String>> sequence,
    bool prefixes = false,
    String column = 'Resistance (Ohms)',
    String swap = 'Ohms',
    List<String> notaion = const [],
  }) {
    int row = 0;
    // Create a title bar in the Excel file.
    for (int i = 0; i < titleBar.length; i++) {
      _file.setCell(index: '${getColumnAlphabet(i)}1', value: titleBar[i].split('=')[0]);
    }

    /// Create a new empty Excel file and generate the title bar in the specified order. Then, open and parse the given Excel files sequentially, extracting the necessary data and inserting it into the corresponding columns.
    /// If a String value in the `titleBar` variable matches a title in the currently open Excel file's title bar, copy all values under that title from the open Excel file and append them to the new Excel file.
    ///
    /// If a String value in the `titleBar` variable contains a single equal sign (`=`), the part before the equal sign is used as a title in the new Excel file, and the part after the equal sign specifies the title in the open Excel file from which the data should be copied.
    /// For example: If the String value is `'Manufacturer Part Number=Mfr Part #'`, the new Excel file will have a title `'Manufacturer Part Number'`, and the corresponding values will be copied from the column titled `'Mfr Part #'` in the open Excel file.
    ///
    /// If a String value in the `titleBar` variable contains two equal signs (`==`), the value following the equal signs will be assigned to every cell in the column corresponding to that title.
    /// For example: Given the value `'Green Certificate1==RoHS Compliant'`, the column titled `'Green Certificate1'` will have `'RoHS Compliant'` assigned to all its cells.
    ///
    /// After the first Excel file that is opened, the first row of each Excel file that is opened is removed and the data is extracted.
    extract() {
      for (String f in files) {
        var temp = LowLevel.open(fileName: f);
        for (int mc = 0; mc < temp.maxCol; mc++) {
          for (int tb = 0; tb < titleBar.length; tb++) {
            // Verify if the String values in the `titleBar` variable match the title bar in the parsed Excel file. If they match, extract the required data from the parsed Excel file accordingly.
            var id = temp.getCell(index: '${getColumnAlphabet(mc)}1');
            if ((id == titleBar[tb]) ||
                ((titleBar[tb].split('=').length > 1) &&
                    (id == titleBar[tb].split('=')[1] ||
                        ('' == titleBar[tb].split('=')[1] && id == titleBar[tb].split('=')[0])))) {
              // Due to having multiple values for the Supplier Part Number in Digikey, only the first Supplier Part Number will be taken as the Supplier Part Number value.
              if (titleBar[tb].split('=')[0] == 'Supplier Part Number 1') {
                for (int i = 2; i <= temp.maxRow; i++) {
                  _file.setCell(
                      index: getColumnAlphabet(tb) + (row + i).toString(),
                      value:
                          temp.getCell(index: getColumnAlphabet(mc) + i.toString())!.split(',')[0]);
                }
              } else {
                // Extracting the required data from the Excel file.
                for (int i = 2; i <= temp.maxRow; i++) {
                  _file.setCell(
                      index: getColumnAlphabet(tb) + (row + i).toString(),
                      value: temp.getCell(index: getColumnAlphabet(mc) + i.toString()) ?? '');
                }
              }
            } else if (titleBar[tb].split('=').length == 3) {
              /**
               * If a String value in the `titleBar` variable contains two equal signs (`==`), the value following the equal signs will be assigned to every cell in the column corresponding to that title.
               * 
               * For example:
               * Given the value `'Green Certificate1==RoHS Compliant'`, the column titled `'Green Certificate1'` will have `'RoHS Compliant'` assigned to all its cells.
               */
              for (int i = 2; i <= temp.maxRow; i++) {
                _file.setCell(
                    index: getColumnAlphabet(tb) + (row + i).toString(),
                    value: titleBar[tb].split('=')[2]);
              }
            }
          }
        }
        row += temp.maxRow - 1;
      }
    }

    /// Used to add characters after a numeric value or to change characters after a numeric value.
    ///
    /// Examples:
    /// 1. If a numerical value like 10.1 is present in the given column, add the string value provided by the swap variable after it. If the value of swap is 'Ω', the new value created will be 10.1Ω.
    /// 2. If values like 10.1Ohm, 11.7KOhm are present in the given column and need to be converted to 10.1Ω, 11.7KΩ respectively, provide the current value and the value it should be changed to in the notation variable. The notation array will be as follows: -> ['Ohm', 'Ω', 'KOhm', 'KΩ']
    prefix({required String column, required String swap, required List<String> notaion}) {
      for (int col = 0; col < _file.maxCol; col++) {
        if (_file.getCell(index: '${getColumnAlphabet(col)}1') == column) {
          for (int row = 1; row <= _file.maxRow; row++) {
            var value = _file.getCell(index: getColumnAlphabet(col) + row.toString());
            if (double.tryParse(value!) != null) {
              _file.setCell(index: getColumnAlphabet(col) + row.toString(), value: value + swap);
            } else {
              int len = (notaion.length % 2 == 0) ? notaion.length : notaion.length - 1;
              for (int i = 0; i < len; i += 2) {
                if (value != value.replaceAll(notaion[i], notaion[i + 1])) {
                  _file.setCell(
                      index: getColumnAlphabet(col) + row.toString(),
                      value: value.replaceAll(notaion[i], notaion[i + 1]));
                }
              }
            }
          }
        }
      }
    }

    // Arrange the data in the given order.(sequence)
    sequenceOfData() {
      List order = List<int>.filled((_file.maxRow * 2) - 2, 0),
          notation = List<int>.filled(_file.maxRow - 1, 0, growable: true),
          numberOfSameValue = List<int>.filled(0, 0, growable: true);
      int notIndex = 0, entryCount = 0, countN = 0, index = 0;
      // order             : Retention of edited, processed data.
      // notation          : For values with a numeric notation, sort the data in the order given. (resistance values)
      // numberOfSameValue : Keeping track of how much data with the same identifier is on the raw data bus.
      // notIndex          : The location where the data is currently being saved in the 'notation' array.
      // entryCount        : Counting the number of entries on map(sequence) one by one.
      // countN            : Retaining the length of the last saved data segment in the notation array.
      // index             : The location where the data is currently being saved in the 'order' array.

      var batchCount, // The number of batches in which the data was previously batches.
          rowLen, // The length of the data bus currently being edited.
          rowStart, // Where to start reading data on the data bus to be edited.
          readRow, // Note the value of the row currently being read.
          readPosition, // Determines whether the processed data in the 'order' array should be read from the first group, or from the second group.
          writePosition, // Determines whether the processed data should be written to the 'order' array in the first group, or in the second group.
          pos, // Saving the length of the data bus in those groups as only the internal data in the preset data groups should be processed during the second data alignment.
          count, // Increment one by one where data should be saved in the 'order' array.
          isCategory; // Whether there was data related to the given category or not.

      // Arrange the data in the given order.
      arrangement() {
        for (var item in sequence.entries) {
          print('Referral Key :- ${item.key}.....................');
          for (int title = 0; title < titleBar.length; title++) {
            if (item.key == titleBar[title].split('=')[0]) {
              // Assigning numberOfSameValue to variables.
              pos = count = 0;
              isCategory = false;
              if (entryCount % 2 == 0) {
                readPosition = _file.maxRow - 1;
                writePosition = 0;
              } else {
                readPosition = 0;
                writePosition = _file.maxRow - 1;
              }
              if (item.key == sequence.keys.first) {
                batchCount = 1;
                rowStart = 1;
                writePosition = 0;
              } else {
                batchCount = numberOfSameValue.length;
                rowStart = 0;
              }
              // Automatically arrange data in alphabetical order.
              if (item.value.first == 'Special_Key') {
                late var first, sta, end;
                // first : Storing the previous data to be compared to sort the data by size.
                // sta   : Where to start saving data when saving data.
                // end   : Where to end saving data when saving data.
                for (var i = 0; i < batchCount; i++) {
                  rowLen =
                      (item.key == sequence.keys.first) ? _file.maxRow + 1 : numberOfSameValue[i];
                  for (var value = 1; value < item.value.length; value++) {
                    // Sorting the raw data using notation.
                    for (var row = rowStart; row < rowLen; row++) {
                      readRow =
                          (item.key == sequence.keys.first) ? row : order[readPosition + pos + row];
                      if (item.value.elementAt(value) ==
                          RegExp(r'[a-zA-Z]+').stringMatch(_file.getCell(
                              index: getColumnAlphabet(title) + readRow.toString())!)) {
                        notation[notIndex++] = readRow;
                      }
                    }
                    // Aligned data is again sorted by size using notation.
                    count = 0;
                    if (notation[0] != 0) {
                      do {
                        first = notation[0];
                        for (var i = 0; i < notIndex - count; i++) {
                          if (notation[i] != -1) {
                            if (double.parse(RegExp(r'(\d*)(\.)*(\d+)')
                                    .stringMatch(_file.getCell(
                                        index: getColumnAlphabet(title) + first.toString())!)
                                    .toString()) >
                                double.parse(RegExp(r'(\d*)(\.)*(\d+)')
                                    .stringMatch(_file.getCell(
                                        index: getColumnAlphabet(title) + notation[i].toString())!)
                                    .toString())) {
                              first = notation[i];
                              i = 0;
                            }
                          }
                        }
                        order[writePosition + countN + count++] = first;
                        notation.removeAt(notation.indexOf(first));
                      } while (count < notIndex);
                      countN += notIndex;
                      notIndex = 0;
                    }
                  }
                  if (item.key != sequence.keys.first) {
                    pos = pos + numberOfSameValue[i] as int;
                  }
                }
                // Saving how much data is available per category.
                count = 1;
                if (entryCount % 2 == 0) {
                  first = order[0];
                  sta = 1;
                  end = _file.maxRow - 1;
                } else {
                  first = order[_file.maxRow - 1];
                  sta = _file.maxRow;
                  end = order.length;
                }
                for (var i = sta; i < end; i++) {
                  if (order[i] != 0) {
                    if (_file.getCell(index: getColumnAlphabet(title) + first.toString()) ==
                        _file.getCell(index: getColumnAlphabet(title) + order[i].toString())) {
                      count++;
                    } else {
                      first = order[i];
                      numberOfSameValue.add(count);
                      count = 1;
                    }
                  }
                }
                numberOfSameValue.add(count);
                if (item.key != sequence.keys.first) {
                  numberOfSameValue.removeRange(0, batchCount);
                }
              } else {
                // Saving data in given order.
                for (int i = 0; i < batchCount; i++) {
                  rowLen =
                      (item.key == sequence.keys.first) ? _file.maxRow + 1 : numberOfSameValue[i];
                  for (int value = 0; value < item.value.length; value++) {
                    for (int row = rowStart; row < rowLen; row++) {
                      readRow =
                          (item.key == sequence.keys.first) ? row : order[readPosition + pos + row];
                      if (item.value.elementAt(value) ==
                          _file.getCell(index: getColumnAlphabet(title) + readRow.toString())!) {
                        order[writePosition + index++] = readRow;
                        isCategory = true;
                      }
                    }
                    if (isCategory) {
                      numberOfSameValue.add(index - count);
                      count = index;
                      isCategory = false;
                    }
                  }
                  if (item.key != sequence.keys.first) {
                    pos = pos + numberOfSameValue[i] as int;
                  }
                }
                index = 0;
                if (item.key != sequence.keys.first) {
                  numberOfSameValue.removeRange(0, batchCount);
                }
              }
            }
          }
          entryCount++;
        }
      }

      // Move the prepared data to a new Excel variable and save it again in the previous Excel variable according to the new order.
      saveData() {
        entryCount--;
        pos = LowLevel.createNew();
        // Save the Heading line.
        for (var col = 0; col < _file.maxCol; col++) {
          pos.setCell(
              index: '${getColumnAlphabet(col)} 1',
              value: _file.getCell(index: '${getColumnAlphabet(col)}1')!);
        }
        if (entryCount % 2 == 0) {
          rowStart = 0;
          rowLen = _file.maxRow - 1;
        } else {
          rowStart = _file.maxRow - 1;
          rowLen = order.length;
        }
        // Processing data by inspecting the Array.
        for (var row = rowStart; row < rowLen; row++) {
          readRow = (entryCount % 2 == 0) ? row + 2 : row + 3 - _file.maxRow;
          for (var col = 0; col < _file.maxCol; col++) {
            if (order[row] != 0) {
              pos.setCell(
                  index: getColumnAlphabet(col) + readRow.toString(),
                  value:
                      _file.getCell(index: getColumnAlphabet(col) + order[row].toString()) ?? '');
            } else {
              pos.setCell(index: getColumnAlphabet(col) + readRow.toString(), value: '');
            }
          }
        }
        // Bringing set data from temporary variable to permanent variable.
        for (var col = 0; col < pos.maxCol; col++) {
          for (var row = 1; row <= pos.maxRow; row++) {
            _file.setCell(
                index: getColumnAlphabet(col) + row.toString(),
                value: pos.getCell(index: getColumnAlphabet(col) + row.toString()) ?? '');
          }
        }
      }

      arrangement();
      saveData();
    }

    extract();
    // if prefixes true, then call prefix().
    if (prefixes) {
      prefix(column: column, swap: swap, notaion: notaion);
    }
    // Calling sequenceOfData() only if a sequence of data is given.
    if (sequence.isNotEmpty) {
      print('Arranging Data.....');
      sequenceOfData();
      print('Data arrangement is successfully finished..');
    }
  }

  /// Examining a particular column and inserting the given data into the other columns given according to the categories of data in it.
  /// check     -> Column to check.
  /// dataList  -> The order in which the data should be entered.
  ///              ['Name of column to enter data_1' : {'Specific data group_1', 'Data to be entered_1', 'Specific data group_2', 'Data to be entered_2'},
  ///               'Name of column to enter data_2' : {'Specific data group_1', 'Data to be entered_1', 'Specific data group_2', 'Data to be entered_2'}]
  void separateData({required String check, required Map<String, List<String>> dataList}) {
    for (int col = 0; col < _file.maxCol; col++) {
      if (_file.getCell(index: '${getColumnAlphabet(col)}1') == check) {
        for (var map in dataList.entries) {
          for (int cols = 0; cols < _file.maxCol; cols++) {
            if (map.key == _file.getCell(index: '${getColumnAlphabet(cols)}1')) {
              int len = (map.value.length % 2 == 0) ? map.value.length : map.value.length - 1;

              for (int row = 2; row <= _file.maxRow; row++) {
                for (int i = 0; i < len; i += 2) {
                  if (_file.getCell(index: getColumnAlphabet(col) + row.toString()) ==
                      map.value[i]) {
                    _file.setCell(
                        index: getColumnAlphabet(cols) + row.toString(), value: map.value[i + 1]);
                  }
                }
              }
            }
          }
        }
        break;
      }
    }
  }

  /// Removes the given special characters from all rows in a given column. [This is to fix an error when converting data from csv to excel.]
  ///
  /// dataList  -> Names of columns to be checked.
  ///
  /// replace   -> The character or string of words to remove.
  void indexRemove({required List<String> dataList, String find = 'Â', String replace = ''}) {
    for (int col = 0; col < _file.maxCol; col++) {
      for (String i in dataList) {
        if (_file.getCell(index: '${getColumnAlphabet(col)}1') == i) {
          for (int row = 2; row <= _file.maxRow; row++) {
            _file.setCell(
                index: getColumnAlphabet(col) + row.toString(),
                value: (_file.getCell(index: getColumnAlphabet(col) + row.toString()) ?? '')
                    .replaceAll(find, replace));
          }
        }
      }
    }
  }

  /// Removing rows with empty cells and, removing additional rows with the same part number.
  void remove({String column = 'Manufacturer Part Number'}) {
    print('Finding Empty Cells and removing....');
    var pos = LowLevel.createNew(); // Create a new temporary Excel file.
    List<int> index = List<int>.filled(0, 0, growable: true);
    // If an empty cell is found, the same row is removed.
    int? colId;
    for (var row = 2; row <= _file.maxRow; row++) {
      for (var col = 0; col < _file.maxCol; col++) {
        if (_file.getCell(index: '${getColumnAlphabet(col)}1') == column) {
          colId = col;
        }
        if ((_file.getCell(index: getColumnAlphabet(col) + row.toString()) == '') ||
            _file.getCell(index: getColumnAlphabet(col) + row.toString()) == null) {
          index.add(row);
          print(
              'An empty cell was encountered. Cell${getColumnAlphabet(col) + row.toString()}.....\nThe entire $row row is removing.');
          break;
        }
      }
    }
    print('Empty Cells finding process is finished');
    // If the same part number is found, all additional rows with the same part number are removed.
    if (colId != null) {
      print('Now finding the same manufacturer part number......');
      var row = 2;
      bool run = true;
      do {
        print('Currently checking line number $row');
        for (var i = 0; i < index.length; i++) {
          if (index[i] == row) {
            run = false;
            break;
          }
        }
        if (run) {
          String? part = _file.getCell(index: getColumnAlphabet(colId) + row.toString());
          for (var i = row + 1; i <= _file.maxRow; i++) {
            if (part == _file.getCell(index: getColumnAlphabet(colId) + i.toString())) {
              index.add(i);
              print('Found the another same part number $i.....\nThe entire $i row is removing.');
            }
          }
        }
        run = true;
        row++;
      } while (row <= _file.maxRow);
    } else {
      print('Warning ⚠: The $column column could not be found.');
    }
    // After the process is finished, bring the file back to the first variable.
    bool type = true;
    int id = 1;
    for (var row = 1; row <= _file.maxRow; row++) {
      for (var i = 0; i < index.length; i++) {
        if (row == index[i]) {
          type = false;
          break;
        }
      }
      if (type) {
        for (var col = 0; col < _file.maxCol; col++) {
          pos.setCell(
              index: getColumnAlphabet(col) + id.toString(),
              value: _file.getCell(index: getColumnAlphabet(col) + row.toString())!);
        }
        id++;
      }
      type = true;
    }
    _file = pos;
  }

  /// Saving the edited excel file.
  void save({path = 'Output', name = 'final'}) => _file.saveFile(path: path, name: name);

  /// Creating descriptions and comments for components.
  ///
  /// details -> Description List of columns from which to extract data for creation.
  ///
  /// index   -> The name of the screen to enter the created description or comment.
  void description({required List<String> details, index = 'Description'}) {
    print('Creating a new  $index......');
    late String discript = '', indexPos;
    List available = List<int>.filled(0, 0, growable: true);
    bool check = true;
    for (var d in details) {
      // Obtaining the position of only currently available columns from the columns from which the data is to be extracted.
      for (var col = 0; col < _file.maxCol; col++) {
        if (d == _file.getCell(index: '${getColumnAlphabet(col)}1')) {
          available.add(col);
        }
        // Get the position of the column to enter the created comment or description.
        if (check && (index == _file.getCell(index: '${getColumnAlphabet(col)}1'))) {
          check = false;
          indexPos = getColumnAlphabet(col);
        }
      }
    }
    // If there is no column in the Excel file to enter the created idea or description, use the last column to enter it and add the required title.
    indexPos = !check ? indexPos : getColumnAlphabet(_file.maxCol);
    _file.setCell(index: '$indexPos 1', value: index);
    for (var row = 2; row <= _file.maxRow; row++) {
      // Extract description from required columns.
      for (var i in available) {
        discript = '$discript ${_file.getCell(index: getColumnAlphabet(i) + row.toString())}';
      }
      // Adding the description to the required column.
      _file.setCell(index: indexPos + row.toString(), value: discript);
      discript = '';
    }
  }

  /// Getting the file list. returning Iterable<String> file names
  ///
  /// path -> main directory path
  Iterable<String> getFiles({required String path}) sync* {
    final dir = Directory(path);
    if (dir.existsSync()) {
      final List<FileSystemEntity> files = dir.listSync();
      for (var f1 in files) {
        print(f1.absolute);
        yield f1.absolute.path.toString();
      }
    } else {
      print('Not Exist');
      yield '';
    }
  }
}

/// A collection of several additional useful functions.
class Additional {
  /// Default save path is D:\\Dart\\altium_library_config\\Info\\temp.txt
  static const String saveAs = 'D:\\Dart\\altium_library_config\\Info\\temp.txt';

  /// The default path of TI packaging data is D:\\Dart\\altium_library_config\\Info\\packageDataOfTI.txt
  static const String tI = 'D:\\Dart\\altium_library_config\\Info\\packageDataOfTI.txt';

  /// Extracting the TI products packaging information using datasheet.
  void packageDataOfTI({String file = tI, String saveAs = saveAs}) {
    List<String> lines = File(file).readAsLinesSync();
    File(saveAs).writeAsStringSync('', flush: true);

    // Column 1
    File(saveAs).writeAsStringSync("'Package Type': [\n", mode: FileMode.append);
    for (String line in lines) {
      String text = "'${line.split(' ')[0]}','${line.split(' ')[2]}',\n";
      File(saveAs).writeAsStringSync(text, mode: FileMode.append);
    }
    File(saveAs).writeAsStringSync('],\n', mode: FileMode.append);

    // Column 2
    File(saveAs).writeAsStringSync("'Package Drawing': [\n", mode: FileMode.append);
    for (String line in lines) {
      String text = "'${line.split(' ')[0]}','${line.split(' ')[3]}',\n";
      File(saveAs).writeAsStringSync(text, mode: FileMode.append);
    }
    File(saveAs).writeAsStringSync('],\n', mode: FileMode.append);

    // Column 3
    File(saveAs).writeAsStringSync("'Lead finish/ Ball material': [\n", mode: FileMode.append);
    for (String line in lines) {
      String text =
          line.split(' Level-1-260C-UNLIM')[0]; // Level-1-260C-UNLIM | Level-2-260C-1 YEAR
      text = "'${line.split(' ')[0]}','${text.split('RoHS & Green ')[1]}',\n";
      File(saveAs).writeAsStringSync(text, mode: FileMode.append);
    }
    File(saveAs).writeAsStringSync('],', mode: FileMode.append);

    print('Packaging Data Extracted');
  }

  /// Extracting titles in given excel file.
  void extractTitle({
    required String file,
    String saveAs = saveAs,
    String listName = 'columnOrder',
    bool flush = true,
  }) {
    var excel = LowLevel.open(fileName: file);
    File(saveAs).writeAsStringSync('List<String> $listName = [\n', flush: flush);
    for (int col = 0; col < excel.maxCol; col++) {
      File(saveAs).writeAsStringSync("'${excel.getCell(index: '${getColumnAlphabet(col)}1')}',\n",
          mode: FileMode.append);
    }
    File(saveAs).writeAsStringSync('];', mode: FileMode.append);
    print('${excel.maxCol} Titles are extracted.');
  }
}
