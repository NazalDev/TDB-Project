import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

//function import
import 'package:tric_duta_bangsa/widgets/sidebar/navigation_sidebar.dart';
import 'package:tric_duta_bangsa/widgets/app_bar/app_bar.dart';

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  @override
  Widget build(BuildContext context) => Scaffold(
    backgroundColor: AppTheme.background,
    drawer: NavigationDrawerWidget(),
    appBar: appBarCustom('TRIC DUTA BANGSA', true),
    body: Stack(
      children: [
        Container(
          margin: EdgeInsets.only(top: 16, left: 16, right: 16),
          decoration: BoxDecoration(color: AppTheme.box),
        ),
      ],
    ),
  );
}
