import 'dart:convert'; // For jsonEncode, utf8, and base64Encode
import 'dart:io'; // For file operations
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:path_provider/path_provider.dart';
import 'package:path/path.dart' as p;
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart'
    hide Column, Row, Border; // Avoid conflict with Flutter's Border
import 'package:flutter/foundation.dart' show kIsWeb;

class IjaraGrafikPage extends StatelessWidget {
  final double narxi;
  final double marja;
  final double sugurta;
  final double gai;
  final double gps;
  final double qqs;
  final double komissiya;
  final int muddat;
  final double boshlangichTolov;
  final List<List<dynamic>> tableData;
  final double jamiLizingSummasi;
  final double jamiTolov;

  IjaraGrafikPage({
    required this.narxi,
    required this.marja,
    required this.sugurta,
    required this.gai,
    required this.gps,
    required this.qqs,
    required this.komissiya,
    required this.muddat,
    required this.boshlangichTolov,
    required this.tableData,
    required this.jamiLizingSummasi,
    required this.jamiTolov,
  });

  @override
  Widget build(BuildContext context) {
    final formatCurrency = NumberFormat('#,##0.00', 'en_US');
    final double jamiOldindanTolov = boshlangichTolov + komissiya;
    final double qoldiqSumma = jamiLizingSummasi - boshlangichTolov;
    final double initialPaymentPercentage =
        (boshlangichTolov / jamiLizingSummasi) * 100;

    // Round the initial payment percentage
    final double roundedInitialPaymentPercentage =
        double.parse(initialPaymentPercentage.toStringAsFixed(2));

    return Scaffold(
      appBar: AppBar(
        title:
            Text('Qaytarish Grafik', style: TextStyle(fontFamily: 'Poppins')),
        backgroundColor: Colors.green,
        actions: [
          IconButton(
            icon: Icon(Icons.file_download),
            onPressed: () {
              saveToExcel(
                narxi: narxi,
                boshlangichTolov: boshlangichTolov,
                komissiya: komissiya,
                sugurta: sugurta,
                qqs: qqs,
                jamiLizingSummasi: jamiLizingSummasi,
                jamiTolov: jamiTolov,
                marja: marja,
                muddat: muddat,
                tableData: tableData,
                initialPaymentPercentage: roundedInitialPaymentPercentage,
              );
            },
          ),
        ],
      ),
      body: SingleChildScrollView(
        padding: EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Card(
              elevation: 2.0,
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(8),
              ),
              child: Padding(
                padding: const EdgeInsets.all(16.0),
                child: Column(
                  children: [
                    _buildSummaryRow('Ijara obyekti xarid bahosi', narxi),
                    _buildSummaryRow('Qoldiq summa', qoldiqSumma),
                    _buildSummaryRow(
                        'Boshlang\'ich bo\'nak to\'lovi $roundedInitialPaymentPercentage%',
                        boshlangichTolov),
                    _buildSummaryRow('QQS', qqs),
                    _buildSummaryRow(
                        'Jami ijara obyekti summasi', jamiLizingSummasi),
                    _buildSummaryRowString('Lizing beruvchining marjasi',
                        '${marja * 100}% yillik'),
                    _buildSummaryRowString('Muddati', '$muddat oy'),
                    _buildSummaryRowString('Imtiyozli davri', '0 oy'),
                    _buildSummaryRow('Xaqiqiy Yillik kelib tushish foizi',
                        _calculateAnnualReturnRate()),
                    _buildSummaryRow('Sug\'urta to\'lovi', sugurta),
                    _buildSummaryRow(
                        'NDS', 600000), // Make sure this is correct
                    _buildSummaryRow(
                        'Bir martalik komissiya to\'lovi', komissiya),
                    _buildSummaryRow(
                        'Ro\'yxatga olish uchun GAI yoki Texnazor', gai),
                    _buildSummaryRow('O\'rtacha yillik daromad summasi',
                        _calculateAverageAnnualIncome()),
                    _buildSummaryRow('O\'rtacha oylik daromad summasi',
                        _calculateAverageMonthlyIncome()),
                    _buildSummaryRow(
                        'Yillik daromad ulushi', _calculateAnnualIncomeShare()),
                    _buildSummaryRow('Jami oldindan to\'lov', jamiTolov),
                  ],
                ),
              ),
            ),
            SizedBox(height: 20),
            SingleChildScrollView(
              scrollDirection: Axis.horizontal,
              child: _buildSummaryTable(),
            ),
            SizedBox(height: 20),
            ElevatedButton(
              onPressed: () {
                Navigator.pop(context);
              },
              child: Text('Ortga qaytish',
                  style: TextStyle(fontFamily: 'Poppins')),
              style: ElevatedButton.styleFrom(
                backgroundColor: Colors.green,
                padding: EdgeInsets.symmetric(vertical: 16),
                minimumSize: Size(double.infinity, 50),
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(8),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  // Methods for calculations, building table rows, and saving to Excel
  double _calculateAnnualReturnRate() {
    final totalUstama = tableData.fold(0.0, (sum, item) => sum + item[2]);
    return (totalUstama / (jamiLizingSummasi - boshlangichTolov)) *
        12 /
        muddat *
        100;
  }

  double _calculateAverageAnnualIncome() {
    final totalUstama = tableData.fold(0.0, (sum, item) => sum + item[2]);
    return (totalUstama / muddat) * 12;
  }

  double _calculateAverageMonthlyIncome() {
    return _calculateAverageAnnualIncome() / 12;
  }

  double _calculateAnnualIncomeShare() {
    return (_calculateAverageAnnualIncome() /
            (jamiLizingSummasi - boshlangichTolov)) *
        100;
  }

  // UI Components for table and rows
  Widget _buildSummaryRow(String label, double value) {
    final formatCurrency = NumberFormat('#,##0.00', 'en_US');
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4.0),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Expanded(child: Text(label, style: TextStyle(fontFamily: 'Poppins'))),
          Expanded(
            child: Text(formatCurrency.format(value) + ' uzs',
                style: TextStyle(fontFamily: 'Poppins'),
                textAlign: TextAlign.end),
          ),
        ],
      ),
    );
  }

  Widget _buildSummaryRowString(String label, String value) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4.0),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Expanded(child: Text(label, style: TextStyle(fontFamily: 'Poppins'))),
          Expanded(
            child: Text(value,
                style: TextStyle(fontFamily: 'Poppins'),
                textAlign: TextAlign.end),
          ),
        ],
      ),
    );
  }

  Widget _buildSummaryTable() {
    final formatCurrency = NumberFormat('#,##0.00', 'en_US');
    final today = DateTime.now();
    final dateFormat = DateFormat('dd/MM/yyyy');

    return Table(
      border: TableBorder.all(color: Colors.grey),
      columnWidths: const <int, TableColumnWidth>{
        0: FixedColumnWidth(40.0),
        1: FixedColumnWidth(120.0),
        2: FixedColumnWidth(120.0),
        3: FixedColumnWidth(120.0),
        4: FixedColumnWidth(120.0),
        5: FixedColumnWidth(140.0),
      },
      children: [
        _buildTableHeader(),
        ...tableData
            .asMap()
            .entries
            .map((entry) => _buildTableRow(entry.value,
                _calculatePaymentDate(today, entry.key + 1), dateFormat))
            .toList(),
        _buildTableTotalRow(),
      ],
    );
  }

  TableRow _buildTableHeader() {
    return TableRow(
      decoration: BoxDecoration(color: Colors.grey[300]),
      children: [
        _buildTableCell('№', isHeader: true),
        _buildTableCell('To\'lov kunlari', isHeader: true),
        _buildTableCell('Qoldiq', isHeader: true),
        _buildTableCell('Ustama', isHeader: true),
        _buildTableCell('Asosiy tani', isHeader: true),
        _buildTableCell('Jami to\'lov', isHeader: true),
      ],
    );
  }

  DateTime _calculatePaymentDate(DateTime startDate, int monthIncrement) {
    int day = startDate.day > 20 ? 20 : startDate.day;

    int year = startDate.year + ((startDate.month + monthIncrement - 1) ~/ 12);
    int month = (startDate.month + monthIncrement - 1) % 12 + 1;

    DateTime paymentDate = DateTime(year, month, day);

    while (paymentDate.month != month) {
      paymentDate = paymentDate.subtract(Duration(days: 1));
    }

    return paymentDate;
  }

  TableRow _buildTableRow(
      List<dynamic> data, DateTime paymentDate, DateFormat dateFormat) {
    final formatCurrency = NumberFormat('#,##0.00', 'en_US');
    return TableRow(
      children: [
        _buildTableCell(data[0].toString()),
        _buildTableCell(dateFormat.format(paymentDate)),
        _buildTableCell(formatCurrency.format(data[1])),
        _buildTableCell(formatCurrency.format(data[2])),
        _buildTableCell(formatCurrency.format(data[3])),
        _buildTableCell(formatCurrency.format(data[2] + data[3])),
      ],
    );
  }

  TableRow _buildTableTotalRow() {
    final formatCurrency = NumberFormat('#,##0.00', 'en_US');
    double totalQoldiq = 0;
    double totalUstama = 0;
    double totalAsosiyTani = 0;
    double totalJamiTolov = 0;

    for (var data in tableData) {
      totalQoldiq += data[1];
      totalUstama += data[2];
      totalAsosiyTani += data[3];
      totalJamiTolov += data[2] + data[3];
    }

    return TableRow(
      decoration: BoxDecoration(color: Colors.grey[200]),
      children: [
        _buildTableCell(''),
        _buildTableCell('Jami', isHeader: true),
        _buildTableCell(formatCurrency.format(totalQoldiq), isHeader: true),
        _buildTableCell(formatCurrency.format(totalUstama), isHeader: true),
        _buildTableCell(formatCurrency.format(totalAsosiyTani), isHeader: true),
        _buildTableCell(formatCurrency.format(totalJamiTolov), isHeader: true),
      ],
    );
  }

  Widget _buildTableCell(String text, {bool isHeader = false}) {
    return Container(
      padding: EdgeInsets.all(8.0),
      decoration: BoxDecoration(
        border: Border.symmetric(horizontal: BorderSide(color: Colors.grey)),
        color: isHeader ? Colors.grey[300] : Colors.white,
      ),
      child: Text(
        text,
        textAlign: TextAlign.center,
        style: TextStyle(
          fontWeight: isHeader ? FontWeight.bold : FontWeight.normal,
          fontFamily: 'Poppins',
        ),
      ),
    );
  }

  void saveToExcel({
    required double narxi,
    required double boshlangichTolov,
    required double komissiya,
    required double sugurta,
    required double qqs,
    required double jamiLizingSummasi,
    required double jamiTolov,
    required double marja,
    required int muddat,
    required List<List<dynamic>> tableData,
    required double initialPaymentPercentage,
  }) async {
    final double jamiOldindanTolov = boshlangichTolov + komissiya;
    final double qoldiqSumma = jamiLizingSummasi - boshlangichTolov;
    final today = DateTime.now();
    final dateFormat = DateFormat('dd/MM/yyyy');

    double totalUstama = tableData.fold(0, (sum, item) => sum + item[2]);
    double averageAnnualIncome = (totalUstama / muddat) * 12;
    double averageMonthlyIncome = averageAnnualIncome / 12;
    double annualReturnRate = (totalUstama / qoldiqSumma) * 12 / muddat * 100;
    double annualIncomeShare = (averageAnnualIncome / qoldiqSumma) * 100;

    final Workbook workbook = Workbook();
    final Worksheet sheet = workbook.worksheets[0];

    String numberFormat = '#,##0';

    sheet.getRangeByName('A1:F2').merge();
    sheet.getRangeByName('A1').setText("To'lov jadvali");
    sheet.getRangeByName('A1').cellStyle.bold = true;
    sheet.getRangeByName('A1').cellStyle.fontSize = 16;
    sheet.getRangeByName('A1').cellStyle.hAlign = HAlignType.center;

    final List<List<dynamic>> summaryData = [
      ["Ijara obyekti xarid bahosi", narxi, "sum", "", "green"],
      ["Qoldiq summa", qoldiqSumma, "sum"],
      [
        "Boshlang'ich bo'nak to'lovi $initialPaymentPercentage%",
        boshlangichTolov,
        "sum",
        sugurta,
        "Sug'urta to'lovi"
      ],
      ["QQS", qqs, "sum", 600000, "NDS"],
      [
        "Jami ijara obyekti summasi",
        jamiLizingSummasi,
        "sum",
        komissiya,
        "Bir martalik komissiya to'lovi"
      ],
      [
        "Lizing beruvchining marjasi",
        "${marja * 100}%",
        "yillik",
        3754600,
        "GAI yoki Texnazor"
      ],
      [
        "Muddati",
        "$muddat",
        "oy",
        averageAnnualIncome,
        "O'rtacha yillik daromad summasi"
      ],
      [
        "Imtiyozli davri",
        "0",
        "oy",
        averageMonthlyIncome,
        "O'rtacha oylik daromad summasi"
      ],
      [
        "Xaqiqiy Yillik kelib tushish foizi",
        "${annualReturnRate.toStringAsFixed(2)}%",
        "yillik",
        annualIncomeShare,
        "Yillik daromad ulushi"
      ],
      ["", "", "", jamiTolov, "Jami to'lov"],
    ];

    for (int i = 0; i < summaryData.length; i++) {
      final row = summaryData[i];
      for (int j = 0; j < row.length; j++) {
        final cell = sheet.getRangeByIndex(i + 3, j + 1);
        if (row[j] is String) {
          cell.setText(row[j]);
        } else if (row[j] is double) {
          cell.setNumber(row[j]);
          cell.numberFormat = numberFormat;
        }
        cell.cellStyle.bold = true;

        if (row.length > 4 && j == 1 && row[4] == "green") {
          cell.cellStyle.backColor = '#00FF00';
        }
        if (row.length > 5 && j == 3 && row[5] == "red") {
          cell.cellStyle.fontColor = '#FF0000';
          cell.cellStyle.backColor = '#FFFF00';
        }
      }
      if (row.length == 5) {
        sheet.getRangeByIndex(i + 3, 4, i + 3, 5).merge();
      }
    }

    // Add text to cells F5 through F12
    final List<String> fColumnTexts = [
      "Sug'urta to'lovi",
      "NDS",
      "Bir martalik komissiya to'lovi",
      "GAI yoki Texnazor",
      "O'rtacha yillik daromad summasi",
      "O'rtacha oylik daromad summasi",
      "Yillik daromad ulushi",
      "Jami to'lov"
    ];

    for (int i = 0; i < fColumnTexts.length; i++) {
      final cell =
          sheet.getRangeByIndex(i + 5, 6); // F column starts at index 6
      cell.setText(fColumnTexts[i]);
      cell.cellStyle.bold = true;
    }

    sheet.getRangeByName('D3:F3').merge();
    sheet.getRangeByName('D3').setText('Xarajatlar');
    sheet.getRangeByName('D3').cellStyle.bold = true;
    sheet.getRangeByName('D3').cellStyle.hAlign = HAlignType.center;

    int startRowIndex = 14;
    final headers = [
      '№',
      'To\'lov kunlari',
      'Qoldiq',
      'Ustama',
      'Asosiy tani',
      'Jami to\'lov'
    ];
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.getRangeByIndex(startRowIndex, i + 1);
      cell.setText(headers[i]);
      cell.cellStyle.bold = true;
      cell.cellStyle.backColor = '#FFFF00';
    }

    double totalQoldiq = 0;
    double totalUstamaSum = 0;
    double totalAsosiyTani = 0;
    double totalJamiTolov = 0;

    for (int i = 0; i < tableData.length; i++) {
      int rowIndex = startRowIndex + 1 + i;
      sheet.getRangeByIndex(rowIndex, 1).setNumber(tableData[i][0]);
      sheet
          .getRangeByIndex(rowIndex, 2)
          .setText(dateFormat.format(_calculatePaymentDate(today, i + 1)));
      sheet.getRangeByIndex(rowIndex, 3).setNumber(tableData[i][1]);
      sheet.getRangeByIndex(rowIndex, 3).numberFormat = numberFormat;
      sheet.getRangeByIndex(rowIndex, 4).setNumber(tableData[i][2]);
      sheet.getRangeByIndex(rowIndex, 4).numberFormat = numberFormat;
      sheet.getRangeByIndex(rowIndex, 5).setNumber(tableData[i][3]);
      sheet.getRangeByIndex(rowIndex, 5).numberFormat = numberFormat;
      sheet
          .getRangeByIndex(rowIndex, 6)
          .setNumber(tableData[i][2] + tableData[i][3]);
      sheet.getRangeByIndex(rowIndex, 6).numberFormat = numberFormat;

      totalQoldiq += tableData[i][1];
      totalUstamaSum += tableData[i][2];
      totalAsosiyTani += tableData[i][3];
      totalJamiTolov += tableData[i][2] + tableData[i][3];
    }

    int totalRowIndex = startRowIndex + 1 + tableData.length;
    sheet.getRangeByIndex(totalRowIndex, 1).setText('Jami:');
    sheet.getRangeByIndex(totalRowIndex, 3).setNumber(totalQoldiq);
    sheet.getRangeByIndex(totalRowIndex, 3).numberFormat = numberFormat;
    sheet.getRangeByIndex(totalRowIndex, 4).setNumber(totalUstamaSum);
    sheet.getRangeByIndex(totalRowIndex, 4).numberFormat = numberFormat;
    sheet.getRangeByIndex(totalRowIndex, 5).setNumber(totalAsosiyTani);
    sheet.getRangeByIndex(totalRowIndex, 5).numberFormat = numberFormat;
    sheet.getRangeByIndex(totalRowIndex, 6).setNumber(totalJamiTolov);
    sheet.getRangeByIndex(totalRowIndex, 6).numberFormat = numberFormat;

    for (int i = 1; i <= 6; i++) {
      sheet.getRangeByIndex(totalRowIndex, i).cellStyle.bold = true;
      sheet.getRangeByIndex(totalRowIndex, i).cellStyle.backColor = '#FFFF00';
    }

    sheet.getRangeByName('A1:F$totalRowIndex').cellStyle.borders.all.lineStyle =
        LineStyle.thin;

    for (int i = 1; i <= 6; i++) {
      sheet.autoFitColumn(i);
    }

    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    if (kIsWeb) {
      final content = base64Encode(bytes);
      final anchor = html.AnchorElement(
          href: 'data:application/octet-stream;base64,$content')
        ..setAttribute('download', 'Qaytarish_Grafik.xlsx')
        ..click();
    } else {
      final directory = await getApplicationSupportDirectory();
      final path = p.join(directory.path, 'Qaytarish_Grafik.xlsx');
      final file = File(path);
      await file.writeAsBytes(bytes, flush: true);
      await Process.run('start', <String>['$path'], runInShell: true);
    }
  }
}
