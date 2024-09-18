import 'package:firebase_auth/firebase_auth.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:firebase_core/firebase_core.dart';
import 'package:hamkoroffline/pages/bot.dart';
import 'Login/employee_dashboard_page.dart';
import 'Login/login_screen.dart';
import 'cars/AllProductsPage.dart';
import 'dashbord_page.dart';
import 'models/leasing_calculator.dart';
import 'models/leasing_services_page.dart';
import 'models/meyory_hujjat.dart';
import 'pages/home_page.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();

  if (kIsWeb) {
    await Firebase.initializeApp(
      options: FirebaseOptions(
        apiKey: "AIzaSyBz9nDOm41hz3c9tRz1p-7H6us8Q5J-JaA",
        authDomain: "hamkorlizing1.firebaseapp.com",
        databaseURL:
            "https://hamkorlizing1-default-rtdb.europe-west1.firebasedatabase.app",
        projectId: "hamkorlizing1",
        storageBucket: "hamkorlizing1.appspot.com",
        messagingSenderId: "259213927261",
        appId: "1:259213927261:web:3c04e3b4d34dcd32303eb4",
        measurementId: "G-SVJNMYDGVM",
      ),
    );
  } else {
    await Firebase.initializeApp();
  }

  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Hamkor Lizing',
      theme: ThemeData(
        primarySwatch: Colors.green,
      ),
      initialRoute: '/home',
      routes: {
        '/home': (context) => HomePage(),
        '/employeeDashboard': (context) =>
            EmployeeDashboardPage(user: FirebaseAuth.instance.currentUser!),
        '/login': (context) => LoginScreen(),
        '/calculator': (context) => LeasingOptionsPage(),
        '/products': (context) =>
            AllProductsPage(categoryTitle: 'Barcha Mahsulotlar'),
        '/documents': (context) => MeyoriyHujjatlarPage(),
        '/services': (context) => LeasingServicesPage(),
      },
      // 404 sahifa uchun unknown route
      onUnknownRoute: (settings) {
        return MaterialPageRoute(builder: (context) => NotFoundPage());
      },
    );
  }
}

class NotFoundPage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text('Sahifa topilmadi')),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Text(
              '404 - Bunday sahifa mavjud emas',
              style: TextStyle(fontSize: 24, color: Colors.red),
            ),
            SizedBox(height: 20),
            ElevatedButton(
              onPressed: () {
                Navigator.pushNamed(context, '/home');
              },
              child: Text('Bosh sahifaga qaytish'),
            ),
          ],
        ),
      ),
    );
  }
}
