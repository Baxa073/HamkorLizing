import 'dart:io';
import 'package:flutter/material.dart';
import 'dart:convert';
import 'package:path_provider/path_provider.dart';
import 'package:intl/intl.dart';
import 'package:path/path.dart' as p;
import 'dart:html' as html;
import 'package:syncfusion_flutter_xlsio/xlsio.dart' hide Column, Row, Border;
import 'package:flutter/foundation.dart' show kIsWeb;

class QaytarishGrafikPage extends StatelessWidget {
  final double narxi;
  final double marja;
  final double sugurta;
  final double qqs;
  final double gai;
  final double gps;
  final double komissiya;
  final int muddat;
  final double boshlangichTolov;
  final List<List<dynamic>> tableData;
  final double jamiLizingSummasi;
  final double jamiTolov;
  final String commissionPercentage;

  QaytarishGrafikPage({
    required this.narxi,
    required this.marja,
    required this.sugurta,
    required this.qqs,
    required this.gai,
    required this.gps,
    required this.komissiya,
    required this.muddat,
    required this.boshlangichTolov,
    required this.tableData,
    required this.jamiLizingSummasi,
    required this.jamiTolov,
    required this.commissionPercentage,
  });

  final formatCurrency = NumberFormat('#,##0.00', 'en_US');

  @override
  Widget build(BuildContext context) {
    final double jamiOldindanTolov = boshlangichTolov + komissiya;
    final double qoldiqSumma = jamiLizingSummasi - boshlangichTolov;
    final today = DateTime.now();

    return Scaffold(
      appBar: AppBar(
        title: Text('Qaytarish Grafik'),
        backgroundColor: Colors.green,
        actions: [
          IconButton(
            icon: Icon(Icons.file_download),
            onPressed: () {
              saveToExcel(
                komissiya: komissiya,
                boshlangichTolov: boshlangichTolov,
                muddat: muddat,
                marja: marja,
                sugurta: sugurta,
                gai: gai,
                gps: gps,
                commissionPercentage: commissionPercentage,
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
            SingleChildScrollView(
              scrollDirection: Axis.horizontal,
              child: Table(
                border: TableBorder.all(color: Colors.grey),
                columnWidths: {
                  0: FixedColumnWidth(150.0),
                  1: FixedColumnWidth(150.0),
                },
                children: [
                  TableRow(
                    decoration: BoxDecoration(color: Colors.grey[300]),
                    children: [
                      _buildTableCell('Lizing obyekti', isHeader: true),
                      _buildTableCell('Avtotransport'),
                    ],
                  ),
                  _buildSummaryRow(
                      'Bir martalik komissiyon to\'lov', komissiya),
                  TableRow(
                    decoration: BoxDecoration(color: Colors.grey[300]),
                    children: [
                      _buildTableCell('Lizing muddati (oy)'),
                      _buildTableCell('$muddat oy'),
                    ],
                  ),
                  _buildSummaryRow('Lizing obyekti xarid bahosi', narxi),
                  _buildSummaryRow('Sug\'urta to\'lovi', sugurta),
                  _buildSummaryRow(
                      'Ro\'yxatga olish (GAI yoki Texnadzor)', gai),
                  _buildSummaryRow('GPS uchun to\'lov', gps),
                  _buildSummaryRow('QQS', qqs),
                  _buildSummaryRow(
                      'Jami lizing summasi QQS bilan', jamiLizingSummasi),
                  TableRow(
                    decoration: BoxDecoration(color: Colors.grey[300]),
                    children: [
                      _buildTableCell('Lizing beruvchining marjasi'),
                      _buildTableCell('${(marja * 100).toStringAsFixed(0)}%'),
                    ],
                  ),
                  _buildSummaryRow(
                      'Boshlang\'ich bo\'nak to\'lovi', boshlangichTolov),
                  _buildSummaryRow(
                      'Jami oldindan to\'lov summasi', jamiOldindanTolov),
                  _buildSummaryRow('Qoldiq summa', qoldiqSumma),
                ],
              ),
            ),
            SizedBox(height: 20),
            ConstrainedBox(
              constraints:
                  BoxConstraints(maxWidth: MediaQuery.of(context).size.width),
              child: SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: _buildSummaryTable(today),
              ),
            ),
            SizedBox(height: 20),
            ElevatedButton(
              onPressed: () {
                Navigator.pop(context);
              },
              child: Text('Ortga qaytish'),
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

  Widget _buildTableCell(String text, {bool isHeader = false}) {
    return Container(
      padding: EdgeInsets.all(8.0),
      color: isHeader ? Colors.grey[300] : Colors.white,
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

  TableRow _buildSummaryRow(String label, double value) {
    return TableRow(
      children: [
        _buildTableCell(label),
        _buildTableCell(formatCurrency.format(value) + ' uzs'),
      ],
    );
  }

  Widget _buildSummaryTable(DateTime today) {
    double totalQoldiq = 0;
    double totalUstama = 0;
    double totalQQSbilan = 0;
    double totalQQSsiz = 0;
    double totalJamiTolov = 0;
    final dateFormat = DateFormat('dd/MM/yyyy');

    for (var data in tableData) {
      totalQoldiq += data[1];
      totalUstama += data[2];
      totalQQSbilan += data[3];
      double qqsSizTolov = data[3] / 1.12;
      totalQQSsiz += qqsSizTolov;
      totalJamiTolov += data[2] + data[3];
    }

    return Table(
      border: TableBorder.all(color: Colors.grey),
      columnWidths: const <int, TableColumnWidth>{
        0: FixedColumnWidth(40.0),
        1: FixedColumnWidth(100.0),
        2: FixedColumnWidth(100.0),
        3: FixedColumnWidth(100.0),
        4: FixedColumnWidth(100.0),
        5: FixedColumnWidth(100.0),
        6: FixedColumnWidth(100.0),
      },
      children: [
        _buildTableHeader(),
        ...tableData.asMap().entries.map((entry) {
          int monthOffset = entry.key + 1;
          DateTime paymentDate = _calculatePaymentDate(today, monthOffset);
          return _buildTableRow(entry.value, paymentDate, dateFormat);
        }).toList(),
        _buildTableTotalRow(totalQoldiq, totalUstama, totalQQSbilan,
            totalQQSsiz, totalJamiTolov),
      ],
    );
  }

  TableRow _buildTableHeader() {
    return TableRow(
      decoration: BoxDecoration(color: Colors.grey[300]),
      children: [
        _buildTableCell('N', isHeader: true),
        _buildTableCell('To\'lov kunlari', isHeader: true),
        _buildTableCell('Qoldiq', isHeader: true),
        _buildTableCell('QQS bilan to\'lov miqdori', isHeader: true),
        _buildTableCell('QQS siz to\'lov miqdori', isHeader: true),
        _buildTableCell('Ustama', isHeader: true),
        _buildTableCell('Jami to\'lov', isHeader: true),
      ],
    );
  }

  TableRow _buildTableRow(
      List<dynamic> data, DateTime paymentDate, DateFormat dateFormat) {
    double qqsSizTolov = data[3] / 1.12;
    return TableRow(
      children: [
        _buildTableCell(data[0].toString()),
        _buildTableCell(dateFormat.format(paymentDate)),
        _buildTableCell(formatCurrency.format(data[1])),
        _buildTableCell(formatCurrency.format(data[3])),
        _buildTableCell(formatCurrency.format(qqsSizTolov)),
        _buildTableCell(formatCurrency.format(data[2])),
        _buildTableCell(formatCurrency.format(data[2] + data[3])),
      ],
    );
  }

  TableRow _buildTableTotalRow(double totalQoldiq, double totalUstama,
      double totalQQSbilan, double totalQQSsiz, double totalJamiTolov) {
    return TableRow(
      decoration: BoxDecoration(color: Colors.grey[200]),
      children: [
        _buildTableCell(''),
        _buildTableCell('Jami', isHeader: true),
        _buildTableCell(formatCurrency.format(totalQoldiq), isHeader: true),
        _buildTableCell(formatCurrency.format(totalQQSbilan), isHeader: true),
        _buildTableCell(formatCurrency.format(totalQQSsiz), isHeader: true),
        _buildTableCell(formatCurrency.format(totalUstama), isHeader: true),
        _buildTableCell(formatCurrency.format(totalJamiTolov), isHeader: true),
      ],
    );
  }

  double parsePercentage(String percentage) {
    return double.tryParse(percentage.replaceAll('%', '').trim()) ?? 0.0;
  }

  void saveToExcel({
    required double komissiya,
    required double boshlangichTolov,
    required int muddat,
    required double marja,
    required double sugurta,
    required double gai,
    required double gps,
    required String commissionPercentage,
  }) async {
    final double jamiOldindanTolov = boshlangichTolov + komissiya;
    final double qoldiqSumma = jamiLizingSummasi - boshlangichTolov;
    final today = DateTime.now();
    final dateFormat = DateFormat('dd/MM/yyyy');

    // Calculate the initial payment percentage
    final double initialPaymentPercentage =
        boshlangichTolov / jamiLizingSummasi * 100;

    // Create a new Excel document.
    final Workbook workbook = Workbook();
    final Worksheet sheet = workbook.worksheets[0];

    // Define the number format for cells
    String numberFormat = '#,##0';

    // Add the title "To'lov Jadvali" merged and centered from A1 to G1
    sheet.getRangeByName('A1:G1').merge();
    sheet.getRangeByName('A1').setText("To'lov Jadvali");
    sheet.getRangeByName('A1').cellStyle
      ..bold = true
      ..hAlign = HAlignType.center
      ..vAlign = VAlignType.center;

    // Set Lizing obyekti nomi and Bino-inshoot
    sheet.getRangeByName('A2:E2').merge();
    sheet.getRangeByName('A2').setText('Lizing obyekti nomi:');
    sheet.getRangeByName('A2').cellStyle.bold = true;
    sheet.getRangeByName('F2:G2').merge();
    sheet.getRangeByName('F2').setText('Avtotransport');
    sheet.getRangeByName('F2').cellStyle.bold = true;

    // Add summary data with highlighted cells
    sheet.getRangeByName('A3:E3').merge();
    sheet
        .getRangeByName('A3')
        .setText('Bir martalik komissiyon to\'lov-$commissionPercentage');
    sheet.getRangeByName('A3').cellStyle.bold = true;
    sheet.getRangeByName('F3').setNumber(komissiya);
    sheet.getRangeByName('F3').numberFormat = numberFormat;
    sheet.getRangeByName('F3').cellStyle.backColor = '#FFFF00';
    sheet.getRangeByName('G3').setText('so\'m');
    sheet.getRangeByName('G3').cellStyle.bold = true;

    sheet.getRangeByName('A4:E4').merge();
    sheet.getRangeByName('A4').setText('Lizing muddati:');
    sheet.getRangeByName('A4').cellStyle.bold = true;
    sheet.getRangeByName('F4').setText('$muddat oy');
    sheet.getRangeByName('F4').cellStyle.bold = true;
    sheet.getRangeByName('G4').setText('oy');
    sheet.getRangeByName('G4').cellStyle.bold = true;

    sheet.getRangeByName('A5:E5').merge();
    sheet.getRangeByName('A5').setText('Lizing obyekti xarid bahosi:');
    sheet.getRangeByName('A5').cellStyle.bold = true;
    sheet.getRangeByName('F5').setNumber(narxi);
    sheet.getRangeByName('F5').numberFormat = numberFormat;
    sheet.getRangeByName('G5').setText('so\'m');
    sheet.getRangeByName('G5').cellStyle.bold = true;

    sheet.getRangeByName('A6:E6').merge();
    sheet.getRangeByName('A6').setText('Sug\'urta to\'lovi uchun (3 yillik):');
    sheet.getRangeByName('A6').cellStyle.bold = true;
    sheet.getRangeByName('F6').setNumber(sugurta);
    sheet.getRangeByName('F6').numberFormat = numberFormat;
    sheet.getRangeByName('G6').setText('so\'m');
    sheet.getRangeByName('G6').cellStyle.bold = true;

    sheet.getRangeByName('A7:E7').merge();
    sheet
        .getRangeByName('A7')
        .setText('Ro\'yxatga olish (GAI yoki Texnadzor):');
    sheet.getRangeByName('A7').cellStyle.bold = true;
    sheet.getRangeByName('F7').setNumber(gai);
    sheet.getRangeByName('F7').numberFormat = numberFormat;
    sheet.getRangeByName('G7').setText('so\'m');
    sheet.getRangeByName('G7').cellStyle.bold = true;

    sheet.getRangeByName('A8:E8').merge();
    sheet.getRangeByName('A8').setText('GPS uchun to\'lov:');
    sheet.getRangeByName('A8').cellStyle.bold = true;
    sheet.getRangeByName('F8').setNumber(gps);
    sheet.getRangeByName('F8').numberFormat = numberFormat;
    sheet.getRangeByName('G8').setText('so\'m');
    sheet.getRangeByName('G8').cellStyle.bold = true;

    sheet.getRangeByName('A9:E9').merge();
    sheet.getRangeByName('A9').setText('QQS:');
    sheet.getRangeByName('A9').cellStyle.bold = true;
    sheet.getRangeByName('F9').setNumber(qqs);
    sheet.getRangeByName('F9').numberFormat = numberFormat;
    sheet.getRangeByName('G9').setText('so\'m');
    sheet.getRangeByName('G9').cellStyle.bold = true;

    sheet.getRangeByName('A10:E10').merge();
    sheet.getRangeByName('A10').setText('Jami lizing summasi QQS bilan:');
    sheet.getRangeByName('A10').cellStyle.bold = true;
    sheet.getRangeByName('F10').setNumber(jamiLizingSummasi);
    sheet.getRangeByName('F10').numberFormat = numberFormat;
    sheet.getRangeByName('F10').cellStyle.backColor = '#ADD8E6';
    sheet.getRangeByName('G10').setText('so\'m');
    sheet.getRangeByName('G10').cellStyle.bold = true;

    sheet.getRangeByName('A11:E11').merge();
    sheet.getRangeByName('A11').setText('Lizing beruvchining marjasi:');
    sheet.getRangeByName('A11').cellStyle.bold = true;
    sheet.getRangeByName('F11').setText('${(marja * 100).toStringAsFixed(0)}%');
    sheet.getRangeByName('F11').cellStyle.bold = true;

    sheet.getRangeByName('A12:E12').merge();
    sheet.getRangeByName('A12').setText(
        'Mijozning oldindan bo\'nak to\'lovi (${initialPaymentPercentage.toStringAsFixed(0)}%):');
    sheet.getRangeByName('A12').cellStyle.bold = true;
    sheet.getRangeByName('F12').setNumber(boshlangichTolov);
    sheet.getRangeByName('F12').numberFormat = numberFormat;
    sheet.getRangeByName('F12').cellStyle.backColor = '#FFFF00';
    sheet.getRangeByName('G12').setText('so\'m');
    sheet.getRangeByName('G12').cellStyle.bold = true;

    sheet.getRangeByName('A13:E13').merge();
    sheet.getRangeByName('A13').setText('Jami oldindan to\'lov summasi :');
    sheet.getRangeByName('A13').cellStyle.bold = true;
    sheet.getRangeByName('F13').setNumber(jamiOldindanTolov);
    sheet.getRangeByName('F13').numberFormat = numberFormat;
    sheet.getRangeByName('F13').cellStyle.backColor = '#FFFF00';
    sheet.getRangeByName('G13').setText('so\'m');
    sheet.getRangeByName('G13').cellStyle.bold = true;

    // Add headers for tableData
    int startRowIndex = 15;
    sheet.getRangeByName('A$startRowIndex').setText('N');
    sheet.getRangeByName('A$startRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('B$startRowIndex').setText('To\'lov kunlari');
    sheet.getRangeByName('B$startRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('C$startRowIndex').setText('Qoldiq');
    sheet.getRangeByName('C$startRowIndex').cellStyle.bold = true;
    sheet
        .getRangeByName('D$startRowIndex')
        .setText('QQS bilan to\'lov miqdori');
    sheet.getRangeByName('D$startRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('E$startRowIndex').setText('QQS siz to\'lov miqdori');
    sheet.getRangeByName('E$startRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('F$startRowIndex').setText('Lizing foizi');
    sheet.getRangeByName('F$startRowIndex').cellStyle.bold = true;
    sheet
        .getRangeByName('G$startRowIndex')
        .setText('Jami lizing oylik to\'lovi');
    sheet.getRangeByName('G$startRowIndex').cellStyle.bold = true;

    // Add data from tableData
    double totalQoldiq = 0;
    double totalQQSbilan = 0;
    double totalQQSsiz = 0;
    double totalUstama = 0;
    double totalJamiTolov = 0;

    for (int i = 0; i < tableData.length; i++) {
      int rowIndex = startRowIndex + 1 + i;
      DateTime paymentDate = _calculatePaymentDate(today, i + 1);
      sheet.getRangeByName('A$rowIndex').setNumber(i + 1);
      sheet
          .getRangeByName('B$rowIndex')
          .setText(dateFormat.format(paymentDate));
      sheet.getRangeByName('C$rowIndex').setNumber(tableData[i][1]);
      sheet.getRangeByName('C$rowIndex').numberFormat = numberFormat;
      sheet.getRangeByName('D$rowIndex').setNumber(tableData[i][3]);
      sheet.getRangeByName('D$rowIndex').numberFormat = numberFormat;
      sheet.getRangeByName('E$rowIndex').setNumber(tableData[i][3] / 1.12);
      sheet.getRangeByName('E$rowIndex').numberFormat = numberFormat;
      sheet.getRangeByName('F$rowIndex').setNumber(tableData[i][2]);
      sheet.getRangeByName('F$rowIndex').numberFormat = numberFormat;
      sheet
          .getRangeByName('G$rowIndex')
          .setNumber(tableData[i][2] + tableData[i][3]);
      sheet.getRangeByName('G$rowIndex').numberFormat = numberFormat;

      totalQoldiq += tableData[i][1];
      totalQQSbilan += tableData[i][3];
      totalQQSsiz += tableData[i][3] / 1.12;
      totalUstama += tableData[i][2];
      totalJamiTolov += tableData[i][2] + tableData[i][3];
    }

    // Add the total row
    int totalRowIndex = startRowIndex + 1 + tableData.length;
    sheet.getRangeByName('A$totalRowIndex').setText('Jami');
    sheet.getRangeByName('A$totalRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('C$totalRowIndex').setNumber(totalQoldiq);
    sheet.getRangeByName('C$totalRowIndex').numberFormat = numberFormat;
    sheet.getRangeByName('C$totalRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('D$totalRowIndex').setNumber(totalQQSbilan);
    sheet.getRangeByName('D$totalRowIndex').numberFormat = numberFormat;
    sheet.getRangeByName('D$totalRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('E$totalRowIndex').setNumber(totalQQSsiz);
    sheet.getRangeByName('E$totalRowIndex').numberFormat = numberFormat;
    sheet.getRangeByName('E$totalRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('F$totalRowIndex').setNumber(totalUstama);
    sheet.getRangeByName('F$totalRowIndex').numberFormat = numberFormat;
    sheet.getRangeByName('F$totalRowIndex').cellStyle.bold = true;
    sheet.getRangeByName('G$totalRowIndex').setNumber(totalJamiTolov);
    sheet.getRangeByName('G$totalRowIndex').numberFormat = numberFormat;
    sheet.getRangeByName('G$totalRowIndex').cellStyle.bold = true;

    // Set all borders
    final Range allDataRange = sheet.getRangeByName('A1:G$totalRowIndex');
    allDataRange.cellStyle.borders.all.lineStyle = LineStyle.thin;

    // Auto-fit columns to make sure all data is visible
    for (int i = 1; i <= 7; i++) {
      sheet.autoFitColumn(i);
    }

    // Save the document.
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
}
