import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'package:tric_duta_bangsa/widgets/sidebar/navigation_provider.dart';
import 'package:tric_duta_bangsa/screens/home_screen.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});
  static const String title = 'TDB APP';

  @override
  Widget build(BuildContext context) => ChangeNotifierProvider(
    create: (context) => NavigationProvider(),
    child: MaterialApp(
      debugShowCheckedModeBanner: false,
      title: title,
      home: HomePage(),
    ),
  );
}
