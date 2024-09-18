import 'package:flutter/material.dart';
import 'morePage/company_info_page.dart';
import 'morePage/company_info_page.dart';

class MorePage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Ko\'proq', style: TextStyle(fontFamily: 'Poppins')),
        backgroundColor: Colors.green,
      ),
      body: ListView(
        padding: EdgeInsets.all(16.0),
        children: <Widget>[
          ListTile(
            leading: Icon(Icons.login, color: Colors.green),
            title: Text('Kirish', style: TextStyle(fontFamily: 'Poppins')),
            trailing: Icon(Icons.arrow_forward, color: Colors.green),
            onTap: () {
              // Kirish sahifasiga o'tish uchun kod
            },
          ),
          ListTile(
            leading: Icon(Icons.info, color: Colors.green),
            title: Text('Kompaniya haqida',
                style: TextStyle(fontFamily: 'Poppins')),
            trailing: Icon(Icons.arrow_forward, color: Colors.green),
            onTap: () {
              Navigator.push(
                context,
                MaterialPageRoute(builder: (context) => CompanyInfoPage()),
              );
            },
          ),
          ListTile(
            leading: Icon(Icons.location_city, color: Colors.green),
            title: Text('Kompaniya filiali',
                style: TextStyle(fontFamily: 'Poppins')),
            trailing: Icon(Icons.arrow_forward, color: Colors.green),
            onTap: () {
              // Kompaniya filiali sahifasiga o'tish uchun kod
            },
          ),
          ListTile(
            leading: Icon(Icons.help, color: Colors.green),
            title: Text('Yordam', style: TextStyle(fontFamily: 'Poppins')),
            trailing: Icon(Icons.arrow_forward, color: Colors.green),
            onTap: () {
              // Yordam sahifasiga o'tish uchun kod
            },
          ),
        ],
      ),
    );
  }
}
