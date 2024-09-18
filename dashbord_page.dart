import 'dart:convert';
import 'dart:html' as html;
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:firebase_auth/firebase_auth.dart'; // Import FirebaseAuth
import 'package:firebase_database/firebase_database.dart';
import 'package:flutter_spinkit/flutter_spinkit.dart';
import 'package:connectivity_plus/connectivity_plus.dart';
import 'package:intl/intl.dart';
import 'package:path_provider/path_provider.dart';
import 'package:pull_to_refresh/pull_to_refresh.dart';
import 'dart:io';
import 'package:lottie/lottie.dart';
import 'package:getwidget/getwidget.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' hide Column, Row;
import '../Login/login_screen.dart'; // Import LoginScreen

class AdminDashboardPage extends StatefulWidget {
  @override
  _AdminDashboardPageState createState() => _AdminDashboardPageState();
}

class _AdminDashboardPageState extends State<AdminDashboardPage> {
  final DatabaseReference databaseReference = FirebaseDatabase.instance.ref();
  List<Map<dynamic, dynamic>> leasingDataList = [];
  List<Map<dynamic, dynamic>> filteredList = [];
  bool isLoading = true;
  bool hasError = false;
  final RefreshController _refreshController =
  RefreshController(initialRefresh: false);

  String selectedRegion = 'Барча';
  String selectedSpecialist = 'Барча';
  bool showAllProjects = true;
  int totalClients = 0;
  int totalDebtors = 0;

  @override
  void initState() {
    super.initState();
    checkInternetConnection();
    fetchData();
  }

  void checkInternetConnection() async {
    var connectivityResult = await (Connectivity().checkConnectivity());
    if (connectivityResult == ConnectivityResult.none) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('Internet bilan aloqa yo‘q'),
          backgroundColor: Colors.red,
        ),
      );
    }
  }

  void fetchData() async {
    try {
      final data =
          (await databaseReference.child("leasing").once()).snapshot.value;
      if (data != null) {
        List<Map<dynamic, dynamic>> tempList = [];
        if (data is Map) {
          data.forEach((key, value) {
            if (value != null) {
              tempList.add(value as Map<dynamic, dynamic>);
            }
          });
        } else if (data is List) {
          data.forEach((value) {
            if (value != null) {
              tempList.add(value as Map<dynamic, dynamic>);
            }
          });
        }
        setState(() {
          leasingDataList = tempList;
          filterData();
          totalClients = leasingDataList.length;
          totalDebtors = leasingDataList.where((client) {
            double debt = double.tryParse(
                client['МУДДАТИ УТГАН ЖАМИ КАРЗ']?.toString() ?? '0') ??
                0;
            return debt > 0;
          }).length;
          isLoading = false;
          hasError = false;
          _refreshController.refreshCompleted();
        });
      } else {
        setState(() {
          isLoading = false;
          hasError = true;
          _refreshController.refreshFailed();
        });
      }
    } catch (error) {
      setState(() {
        hasError = true;
        isLoading = false;
        _refreshController.refreshFailed();
      });
      print('Error fetching data: $error');
    }
  }

  void filterData() {
    setState(() {
      filteredList = leasingDataList.where((client) {
        bool matchesRegion = selectedRegion == 'Барча' ||
            client['Мижоз рўйхатдан ўтган худуд']?.toString() == selectedRegion;
        bool matchesSpecialist = selectedSpecialist == 'Барча' ||
            client['Мутахассиси']?.toString() == selectedSpecialist;

        if (!matchesRegion || !matchesSpecialist) {
          return false;
        }

        if (showAllProjects) {
          return true;
        } else {
          double overdueAmount = double.tryParse(
              client['МУДДАТИ УТГАН ЖАМИ КАРЗ']?.toString() ?? '0') ??
              0;
          return overdueAmount > 0;
        }
      }).toList();
    });
  }

  Future<void> _logout() async {
    await FirebaseAuth.instance.signOut();
    Navigator.pushReplacement(
      context,
      MaterialPageRoute(builder: (context) => LoginScreen()),
    );
  }

  void generateExcel() async {
    try {
      List<Map<dynamic, dynamic>> dataListToSave = showAllProjects
          ? filteredList
          : filteredList.where((client) {
        double overdueAmount = double.tryParse(
            client['МУДДАТИ УТГАН ЖАМИ КАРЗ']?.toString() ?? '0') ??
            0;
        return overdueAmount > 0;
      }).toList();

      final Workbook workbook = Workbook();
      final Worksheet sheet = workbook.worksheets[0];

      _addHeaders(sheet);
      _addDataRows(sheet, dataListToSave);

      _autoFitColumns(sheet);

      final List<int> bytes = workbook.saveAsStream();
      workbook.dispose();

      saveExcelFile(bytes, 'Umumiy_Hisobot.xlsx');
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Excel file generated successfully!')),
      );
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Error generating Excel file: $e')),
      );
    }
  }

  void _addHeaders(Worksheet sheet) {
    List<String> headers = [
      'N',
      'Мижоз номи',
      'Шартнома раками',
      'Лизинг берилган сана',
      'Лизингни сундириш муддати',
      'Лизинг шартномаси суммаси',
      'Лойиҳадаги мижознинг пул кўринишидаги улуши',
      'Дастлабки лизинг колдиги суммаси',
      'Хисобот кунига лизинг колдик суммаси',
      'Лизинг объекти',
      'Муддати утган лизинг танидан',
      'Муддати утган Фоиздан',
      'Пеня',
      'МУДДАТИ УТГАН ЖАМИ КАРЗ',
      'Мижоз рўйхатдан ўтган худуд',
      'Мижоз телефон рақами',
      'Мутахассиси'
    ];

    final Style headerStyle = _createHeaderStyle(sheet);
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.getRangeByIndex(1, i + 1);
      cell.setText(headers[i]);
      cell.cellStyle = headerStyle;
    }
  }

  Style _createHeaderStyle(Worksheet sheet) {
    final Style headerStyle = sheet.workbook.styles.add('HeaderStyle');
    headerStyle.fontName = 'Calibri';
    headerStyle.bold = true;
    headerStyle.hAlign = HAlignType.center;
    headerStyle.vAlign = VAlignType.center;
    headerStyle.borders.all.lineStyle = LineStyle.thin;
    return headerStyle;
  }

  void _addDataRows(Worksheet sheet, List<Map<dynamic, dynamic>> dataListToSave) {
    final Style numberCellStyle = _createNumberCellStyle(sheet);

    for (int row = 0; row < dataListToSave.length; row++) {
      final client = dataListToSave[row];
      sheet.getRangeByIndex(row + 2, 1).setNumber(row + 1); // Counter column

      for (int col = 1; col < 17; col++) {
        final cell = sheet.getRangeByIndex(row + 2, col + 1);
        var dataValue = client[_getColumnKey(col)];

        if (_isNumericColumn(col)) {
          _setNumericValue(cell, dataValue, numberCellStyle);
        } else {
          cell.setText(dataValue?.toString() ?? '');
        }

        cell.cellStyle.borders.all.lineStyle = LineStyle.thin;
      }
    }
  }

  bool _isNumericColumn(int col) {
    return [5, 6, 7, 8, 10, 11, 12, 13].contains(col);
  }

  String _getColumnKey(int col) {
    List<String> keys = [
      'N',
      'Мижоз номи',
      'Шартнома раками',
      'Лизинг берилган сана',
      'Лизингни сундириш муддати',
      'Лизинг шартномаси суммаси',
      'Лойиҳадаги мижознинг пул кўринишидаги улуши',
      'Дастлабки лизинг колдиги суммаси',
      'Хисобот кунига лизинг колдик суммаси',
      'Лизинг объекти',
      'Муддати утган лизинг танидан',
      'Муддати утган Фоиздан',
      'Пеня',
      'МУДДАТИ УТГАН ЖАМИ КАРЗ',
      'Мижоз рўйхатдан ўтган худуд',
      'Мижоз телефон рақами',
      'Мутахассиси'
    ];
    return keys[col];
  }

  void _setNumericValue(Range cell, var dataValue, Style numberCellStyle) {
    if (dataValue is num) {
      cell.setNumber(dataValue.toDouble());
      cell.cellStyle = numberCellStyle;
    } else if (dataValue is String) {
      double? value = double.tryParse(dataValue);
      if (value != null) {
        cell.setNumber(value);
        cell.cellStyle = numberCellStyle;
      } else {
        cell.setText('N/A');
      }
    }
  }

  Style _createNumberCellStyle(Worksheet sheet) {
    final Style numberCellStyle = sheet.workbook.styles.add('NumberCellStyle');
    numberCellStyle.hAlign = HAlignType.right;
    numberCellStyle.vAlign = VAlignType.center;
    numberCellStyle.numberFormat = '#,##0.00';
    numberCellStyle.borders.all.lineStyle = LineStyle.thin;
    return numberCellStyle;
  }

  void _autoFitColumns(Worksheet sheet) {
    for (int i = 1; i <= 17; i++) {
      sheet.autoFitColumn(i);
    }
  }

  void saveExcelFile(List<int> bytes, String fileName) async {
    if (kIsWeb) {
      final content = base64Encode(bytes);
      final anchor = html.AnchorElement(
          href: 'data:application/octet-stream;base64,$content')
        ..setAttribute('download', fileName)
        ..click();
    } else {
      final directory = await getApplicationDocumentsDirectory();
      final path = '${directory.path}/$fileName';
      final file = File(path);
      await file.writeAsBytes(bytes, flush: true);
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('File saved: $path')),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Admin Paneli'),
        backgroundColor: Colors.green,
        actions: [
          IconButton(
            icon: Icon(Icons.save_alt),
            onPressed: generateExcel,
          ),
          IconButton(
            icon: Icon(Icons.logout),
            onPressed: _logout, // Call the logout method
          ),
        ],
      ),
      body: SmartRefresher(
        controller: _refreshController,
        onRefresh: fetchData,
        header: WaterDropHeader(
          waterDropColor: Colors.green,
        ),
        child: Container(
          padding: EdgeInsets.all(16.0),
          decoration: BoxDecoration(
            gradient: LinearGradient(
              colors: [Colors.white, Colors.green[100]!],
              begin: Alignment.topCenter,
              end: Alignment.bottomCenter,
            ),
          ),
          child: isLoading
              ? Center(
            child: SpinKitFadingCircle(
              color: Colors.green,
              size: 50.0,
            ),
          )
              : hasError
              ? Center(
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                Lottie.asset('assets/animations/not_found.json'),
                SizedBox(height: 20),
                Text(
                  'Ma’lumotlar yuklanmadi!',
                  style: TextStyle(fontSize: 18, color: Colors.red),
                ),
              ],
            ),
          )
              : SingleChildScrollView(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                buildSummaryCards(),
                SizedBox(height: 20),
                buildSectionTitle('Xodimlarning Loyihalar soni'),
                buildSpecialistsList(),
                SizedBox(height: 20),
                buildSectionTitle('Xodimlarning Qarzdor Loyihalar soni'),
                buildDebtorsList(),
                SizedBox(height: 20),
                buildFilters(),
                SizedBox(height: 20),
                buildCheckboxOptions(),
                SizedBox(height: 20),
                buildProjectSummary(),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Widget buildSectionTitle(String title) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 8.0),
      child: Text(
        title,
        style: TextStyle(
            fontSize: 20, fontWeight: FontWeight.bold, color: Colors.green[800]),
      ),
    );
  }

  Widget buildSummaryCards() {
    return Row(
      mainAxisAlignment: MainAxisAlignment.spaceBetween,
      children: [
        buildSummaryCard(
            'Jami Mijozlar', totalClients.toString(), Icons.people, Colors.blue),
        buildSummaryCard(
            'Qarzdorlar', totalDebtors.toString(), Icons.error, Colors.red),
      ],
    );
  }

  Widget buildSummaryCard(
      String title, String value, IconData icon, Color color) {
    return Expanded(
      child: Card(
        elevation: 4,
        shadowColor: color.withOpacity(0.5),
        color: color.withOpacity(0.1),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
        child: Padding(
          padding: const EdgeInsets.all(20.0),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Icon(icon, color: color, size: 48),
              SizedBox(height: 10),
              Text(
                value,
                style: TextStyle(
                    fontSize: 28, fontWeight: FontWeight.bold, color: color),
              ),
              Text(title, style: TextStyle(fontSize: 16, color: color)),
            ],
          ),
        ),
      ),
    );
  }

  Widget buildSpecialistsList() {
    final specialistsData = _getSpecialistsData();
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: specialistsData.map((entry) {
        return Card(
          elevation: 2,
          margin: EdgeInsets.symmetric(vertical: 8.0),
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
          child: ListTile(
            leading: CircleAvatar(
              backgroundColor: Colors.green[300],
              child: Icon(Icons.person, color: Colors.white),
            ),
            title: Text(entry.key,
                style: TextStyle(fontWeight: FontWeight.bold)),
            subtitle: Text('Loyiha soni: ${entry.value}'),
          ),
        );
      }).toList(),
    );
  }

  Widget buildDebtorsList() {
    final debtorsData = _getDebtorsData();
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: debtorsData.map((entry) {
        return Card(
          elevation: 2,
          margin: EdgeInsets.symmetric(vertical: 8.0),
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
          child: ListTile(
            leading: CircleAvatar(
              backgroundColor: Colors.red[300],
              child: Icon(Icons.warning, color: Colors.white),
            ),
            title: Text(entry.key,
                style: TextStyle(fontWeight: FontWeight.bold)),
            subtitle: Text('Qarzdor loyihalar: ${entry.value}'),
          ),
        );
      }).toList(),
    );
  }

  List<MapEntry<String, int>> _getSpecialistsData() {
    Map<String, int> specialistsCount = {};
    for (var client in leasingDataList) {
      String specialist =
          client['Мутахассиси']?.toString().toUpperCase() ?? 'НЕИЗВЕСТНО';
      specialistsCount[specialist] =
          (specialistsCount[specialist] ?? 0) + 1;
    }
    return specialistsCount.entries.toList()
      ..sort((a, b) => b.value.compareTo(a.value));
  }

  List<MapEntry<String, int>> _getDebtorsData() {
    Map<String, int> debtorsCount = {};
    for (var client in leasingDataList) {
      String specialist =
          client['Мутахассиси']?.toString().toUpperCase() ?? 'НЕИЗВЕСТНО';
      double debt = double.tryParse(
          client['МУДДАТИ УТГАН ЖАМИ КАРЗ']?.toString() ?? '0') ??
          0;
      if (debt > 0) {
        debtorsCount[specialist] = (debtorsCount[specialist] ?? 0) + 1;
      }
    }
    return debtorsCount.entries.toList()
      ..sort((a, b) => b.value.compareTo(a.value));
  }

  Widget buildFilters() {
    List<String> regions = [
      'Барча',
      'Андижанская',
      'Ташкентская',
      'Бухарская',
      'Ферганская',
      'Джизакская',
      'Навоийская',
      'Кашкадарьинская',
      'Самаркандская',
      'Сурхандарьинская',
      'Хорезмская',
      'Республика Каракалпакистан',
      'г. Ташкент'
    ];
    List<String> specialists = [
      'Барча',
      'ALIKULOV ABDUKARIM ABDUKAMALOVICH',
      'TIRKASHBOYEV NASIBJON KAMOLDIN O‘G‘LI',
      'SAYIDOV ABDULLAXON MUTALIPJONOVICH',
      'TOSHQINOV BAXRUS BUXOROVICH',
      'ERNAZAROV ASQARALI NURMAMАТОВИЧ',
      'SHARIPOV SHOXBOZ NU’MONJONO‘G‘LI',
      'KUZIBAYEV FERUZ OCHILOVICH'
    ];

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        buildDropdown(
          label: "Худудни танланг",
          value: selectedRegion,
          items: regions,
          onChanged: (newValue) {
            setState(() {
              selectedRegion = newValue!;
              filterData();
            });
          },
        ),
        SizedBox(height: 10),
        buildDropdown(
          label: "Мутахасисни танланг",
          value: selectedSpecialist,
          items: specialists,
          onChanged: (newValue) {
            setState(() {
              selectedSpecialist = newValue!;
              filterData();
            });
          },
        ),
      ],
    );
  }

  Widget buildDropdown({
    required String label,
    required String value,
    required List<String> items,
    required ValueChanged<String?> onChanged,
  }) {
    return DropdownButtonFormField<String>(
      decoration: InputDecoration(
        labelText: label,
        border: OutlineInputBorder(
          borderRadius: BorderRadius.circular(10),
        ),
      ),
      value: value,
      items: items.map((String value) {
        return DropdownMenuItem<String>(
          value: value,
          child: Text(value),
        );
      }).toList(),
      onChanged: onChanged,
    );
  }

  Widget buildCheckboxOptions() {
    return Row(
      children: [
        buildCheckbox("Hammasi", showAllProjects, (value) {
          setState(() {
            showAllProjects = value!;
            filterData();
          });
        }),
        SizedBox(width: 20),
        buildCheckbox("Muddati O'tgan", !showAllProjects, (value) {
          setState(() {
            showAllProjects = !value!;
            filterData();
          });
        }),
      ],
    );
  }

  Widget buildCheckbox(
      String title, bool value, ValueChanged<bool?> onChanged) {
    return Row(
      children: [
        Checkbox(
          value: value,
          activeColor: Colors.green,
          onChanged: onChanged,
        ),
        Text(title),
      ],
    );
  }

  Widget buildProjectSummary() {
    return GFCard(
      boxFit: BoxFit.cover,
      titlePosition: GFPosition.start,
      showImage: true,
      title: GFListTile(
        titleText: 'Loyiha hisobotlari',
      ),
      content: SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: DataTable(
          columns: _buildColumns(),
          rows: _buildRows(),
          showCheckboxColumn: false,
          sortAscending: true,
          sortColumnIndex: 0,
        ),
      ),
    );
  }

  List<DataColumn> _buildColumns() {
    return [
      DataColumn(label: Text('т/р')),
      DataColumn(label: Text('Мижоз номи')),
      DataColumn(label: Text('Шартнома раками')),
      DataColumn(label: Text('Лизинг берилган сана')),
      DataColumn(label: Text('Лизингни сундириш муддати')),
      DataColumn(label: Text('Лизинг шартномаси суммаси')),
      DataColumn(
          label: Text('Лойиҳадаги мижознинг пул кўринишидаги улуши')),
      DataColumn(label: Text('Дастлабки лизинг колдиги суммаси')),
      DataColumn(label: Text('Хисобот кунига лизинг колдик суммаси')),
      DataColumn(label: Text('Лизинг объекти')),
      DataColumn(label: Text('Муддати утган лизинг танидан')),
      DataColumn(label: Text('Муддати утган Фоиздан')),
      DataColumn(label: Text('Пеня')),
      DataColumn(label: Text('МУДДАТИ УТГАН ЖАМИ КАРЗ')),
      DataColumn(label: Text('Мижоз рўйхатдан ўтган худуд')),
      DataColumn(label: Text('Мижоз телефон рақами')),
      DataColumn(label: Text('Мутахассиси')),
    ];
  }

  List<DataRow> _buildRows() {
    final NumberFormat currencyFormat =
    NumberFormat.currency(locale: 'en_US', symbol: '', decimalDigits: 2);
    return filteredList.asMap().entries.map((entry) {
      final client = entry.value;
      return DataRow(cells: [
        DataCell(Text((entry.key + 1).toString())), // Counter
        DataCell(Text(client['Мижоз номи'].toString())),
        DataCell(Text(client['Шартнома раками'].toString())),
        DataCell(Text(client['Лизинг берилган сана'].toString())),
        DataCell(Text(client['Лизингни сундириш муддати'].toString())),
        DataCell(Text(currencyFormat.format(double.tryParse(
            client['Лизинг шартномаси суммаси'].toString()) ??
            0))),
        DataCell(Text(currencyFormat.format(double.tryParse(
            client['Лойиҳадаги мижознинг пул кўринишидаги улуши']
                .toString()) ??
            0))),
        DataCell(Text(currencyFormat.format(double.tryParse(
            client['Дастлабки лизинг колдиги суммаси'].toString()) ??
            0))),
        DataCell(Text(currencyFormat.format(double.tryParse(
            client['Хисобот кунига лизинг колдик суммаси']
                .toString()) ??
            0))),
        DataCell(Text(client['Лизинг объекти'].toString())),
        DataCell(Text(currencyFormat.format(double.tryParse(
            client['Муддати утган лизинг танидан'].toString()) ??
            0))),
        DataCell(Text(currencyFormat.format(double.tryParse(
            client['Муддати утган Фоиздан'].toString()) ??
            0))),
        DataCell(Text(currencyFormat
            .format(double.tryParse(client['Пеня'].toString()) ?? 0))),
        DataCell(Text(currencyFormat.format(double.tryParse(
            client['МУДДАТИ УТГАН ЖАМИ КАРЗ'].toString()) ??
            0))),
        DataCell(Text(client['Мижоз рўйхатдан ўтган худуд'].toString())),
        DataCell(Text(client['Мижоз телефон рақами'].toString())),
        DataCell(Text(client['Мутахассиси'].toString())),
      ]);
    }).toList();
  }
}
