import 'dart:io';
import 'content.dart';
import 'package:path/path.dart';

/// The order in which the columns should be in the new Excel file to be created.
/// Ensure that the data for all the columns below is extracted from the Excel file. Otherwise, if the remove() function is called, all the data may be deleted.
/// Data Format [The title for the columns in the new Excel file=Column title if extracting a column from old Excel file=If there is only one data for all the rows of the created column, then that data is included.]
///
/// If a String value in the this variable contains a single equal sign (`=`), the part before the equal sign is used as a title in the new Excel file, and the part after the equal sign specifies the title in the open Excel file from which the data should be copied.
/// For example: If the String value is `'Manufacturer Part Number=Mfr Part #'`, the new Excel file will have a title `'Manufacturer Part Number'`, and the corresponding values will be copied from the column titled `'Mfr Part #'` in the open Excel file.
///
/// If a String value in the this variable contains two equal signs (`==`), the value following the equal signs will be assigned to every cell in the column corresponding to that title.
/// For example: Given the value `'Green Certificate1==RoHS Compliant'`, the column titled `'Green Certificate1'` will have `'RoHS Compliant'` assigned to all its cells.
List<String> columnOrder = [
  'Manufacturer=Mfr',
  'Manufacturer Part Number=Mfr Part #',
  'Supplier 1==Digi-Key',
  'Supplier Part Number 1=DK Part #',
  'Category==Capasitors',
  'Series==TANTAMOUNT®, 593D',
  'Capacitance',
  'Tolerance',
  'Voltage - Rated',
  'ESR (Equivalent Series Resistance)',
  'Operating Temperature',
  'Type',
  'Manufacturer Size Code',
  'Package / Case',
  'Size / Dimension',
  'Height - Seated (Max)',
  'Comment',
  'Description',
  'Ratings==-',
  'Eco Plan==RoHS & Green',
  'Datasheet==https://www.vishay.com/docs/40005/593d.pdf',
  'Library Path==Library Path',
  'Library Ref==Fix_Polarized_Capacitor',
  'Footprint Path==Footprint Path A',
  'Footprint Path 2==Footprint Path B',
  'Footprint Ref',
  'Footprint Ref 2'
];

/// The order in which all components should be aligned.
/// NOTE: When sorting data in ascending order, if the column contains special characters (e.g., µ, Á, ð, ñ), using a REGEX filter may cause errors. Therefore, these special characters must first be converted into simpler characters. For example, µ = u, Á = A. This can be done by setting the value of the prefixes variable to true in the sheetsConfig function and then using the prefixes function. After sorting the column in ascending order, the indexRemove function can be used to retrieve the special characters that were previously converted to simpler characters.
Map<String, Set<String>> sequence = {
  'Manufacturer Size Code': {'A', 'B', 'C', 'D', 'E'},
  'Capacitance': {'Special_Key', 'uF'}
};

/// The order in which the data should be entered for the description.
List<String> descriptions = [
  'Capacitance',
  'Tolerance',
  'Voltage - Rated',
  'ESR (Equivalent Series Resistance)',
  'Lifetime @ Temp.',
  'Operating Temperature',
  'Type',
  'Package / Case',
];

/// The order in which the data should be entered for the comments.
List<String> comments = ['Capacitance'];

/// By checking the given column, according to the data groups in that column, the order in which the data should be entered in the other columns given to it.
Map<String, List<String>> splitOrder = {
  'Footprint Ref': [
    'A',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_A_N',
    'B',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_B_N',
    'C',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_C_N',
    'D',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_D_N',
    'E',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_E_N',
  ],
  'Footprint Ref 2': [
    'A',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_A_L',
    'B',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_B_L',
    'C',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_C_L',
    'D',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_D_L',
    'E',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_E_L',
  ],
  'Footprint Ref 3': [
    'A',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_A_M',
    'B',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_B_M',
    'C',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_C_M',
    'D',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_D_M',
    'E',
    'VISHAY_TANTAMOUNT-293D/593D/TR3_E_M',
  ],
};

/// Names of columns where a word needs to be replaced by another word.
List<String> remove = ['Capacitance'];

/// Used to add characters after a numeric value or to change characters after a numeric value.
///
/// Examples:
/// 1. If a numerical value like 10.1 is present in the given column, add the string value provided by the swap variable after it. If the value of swap is 'Ω', the new value created will be 10.1Ω.
/// 2. If values like 10.1Ohm, 11.7KOhm are present in the given column and need to be converted to 10.1Ω, 11.7KΩ respectively, provide the current value and the value it should be changed to in the notation variable. The notation array will be as follows: -> ['Ohm', 'Ω', 'KOhm', 'KΩ']
List<String> prefixOrder = ['µF', ' uF'];

// TODO: RawData folder contains Excel files that need to be edited and the data is saved to the Output folder after editing. For ease of understanding the code later, there are several Excel files for editing in the RawData folder. So, run the program and see how it works.

void main() {
  // Use the terminal to run the program
  // Terminal command -> dart bin\altium_library_config.dart

  String directory = 'D:\\Dart\\altium_library_config\\RawData\\';
  print('Do you want to change the default file directory path? (Yes or No)');
  if (stdin.readLineSync() == 'Yes') {
    print(
        'Attention Please ⚠️: Please ensure the directory contains only .xlsx files, no files are currently open, and the correct file path is selected. Non-compliance with these conditions will result in an error.\n'
        'Enter the File Directory\n');
    directory = stdin.readLineSync()!;
  } else {
    print('Ok, We will continue to use the default file directory');
  }

  /// This function takes given Excel files and arranges them in the given order.
  marge() {
    try {
      final Stopwatch stopwatch = Stopwatch()..start();
      // Check user given directory and get all filenames in given directory.
      var files = Processing().getFiles(path: directory).toList();
      Processing()
        ..sheetsConfig(
            titleBar: columnOrder,
            files: files,
            sequence: sequence,
            prefixes: true,
            column: 'Capacitance',
            swap: ' uF',
            notaion: prefixOrder)
        // Checking the given column and entering the data in other given columns according to the categories of data in that column.
        ..separateData(check: 'Manufacturer Size Code', dataList: splitOrder)
        // Checks all the rows in a given column and replaces any word in them with another word.
        ..indexRemove(dataList: remove, find: 'uF', replace: 'µF')
        // Generating a description for components.
        ..description(details: descriptions)
        // Generating a comments for components.
        ..description(details: comments, index: 'Comment')
        // Removing multiple columns with the same component and removing columns with empty cells.
        ..remove()
        // Save the file.
        ..save(name: 'test_VISHAY_593D');
      stopwatch.stop();
      print('Elapsed time: ${(stopwatch.elapsedMilliseconds) / 1000} S');
      Process.runSync('powershell', ['-c', '[console]::beep(1000, 1000)']);
      print('Well Done.... Program is Completed');
    } catch (e) {
      print("Something's wrong");
      print(e);
      exit(255);
    }
  }

  /// This function arranges the Excel files that are to be arranged in the same order at the same time according to the given order.
  /// For example, if a column with the same value should be added to the pre-prepared Excel sheets with the same sequence (new column => 'Ratings==AEC-Q200'), it can be done by this function.
  simultaneous() {
    try {
      final Stopwatch stopwatch = Stopwatch()..start();
      // Check user given directory and get all filenames in given directory.
      var files = Processing().getFiles(path: directory).toList();
      for (var f in files) {
        Processing()
          ..sheetsConfig(
              titleBar: columnOrder,
              files: f.split(' '),
              sequence: sequence,
              prefixes: false,
              column: 'Capacitance',
              swap: 'uF',
              notaion: prefixOrder)
          // Checking the given column and entering the data in other given columns according to the categories of data in that column.
          ..separateData(check: 'Manufacturer Size Code', dataList: splitOrder)
          // Checks all the rows in a given column and replaces any word in them with another word.
          ..indexRemove(dataList: remove, find: 'uF', replace: 'µF')
          // Generating a description for components.
          ..description(details: descriptions)
          // Generating a comments for components.
          ..description(details: comments, index: 'Comment')
          // Removing multiple columns with the same component and removing columns with empty cells.
          ..remove()
          // Save the file.
          ..save(name: basename(f).split('.xlsx')[0]);
        print('${basename(f)} is Completed');
      }
      stopwatch.stop();
      print('Elapsed time: ${(stopwatch.elapsedMilliseconds) / 1000} S');
      Process.runSync('powershell', ['-c', '[console]::beep(1000, 1000)']);
      print('Well Done.... Program is Completed');
    } catch (e) {
      print("Something's wrong");
      print(e);
      exit(255);
    }
  }

  marge();

  //  simultaneous();

  //  Additional().packageDataOfTI();
  //  Additional().extractTitle(file: 'D:\\Dart\\altium_library_config\\RawData\\fileName.xlsx'); //  Replace fileName with the name of the required file
}
