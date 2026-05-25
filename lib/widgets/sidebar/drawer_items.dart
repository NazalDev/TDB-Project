import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/widgets/sidebar/drawer_item.dart';

/*
sideBarTile(() {}, 'Home', Icons.home),
              sideBarTile(() {}, 'Dashboard', Icons.dashboard),
              sideBarTile(() {}, 'Laporan', Icons.report),

              Divider(
                color: AppTheme.greenDark,
                thickness: 5,
                radius: BorderRadius.circular(8),
              ),

              // Supply Chain Management
              sideBarTile(() {}, 'PO & DO Page', Icons.file_upload),
              sideBarTile(() {}, 'Product', Icons.archive),
              sideBarTile(() {}, 'Invoice', Icons.download),
              ListTile(),
*/

/*
  Padding sideBarTile(Function()? onTap, String title, IconData icon) {
    return Padding(
      padding: const EdgeInsets.only(top: 4, bottom: 4),
      child: ListTile(
        hoverColor: AppTheme.hover,
        splashColor: AppTheme.clicked,
        onTap: onTap,
        leading: Icon(icon, color: AppTheme.greenDark),
        title: Text(
          title,
          style: TextStyle(
            fontFamily: 'Montserrat',
            fontWeight: FontWeight.bold,
            color: AppTheme.textPrimary,
          ),
        ),
        tileColor: AppTheme.buttonNormal,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
      ),
    );
  }
*/

final itemsFirst = [
  DrawerItem(title: 'Home', icon: Icons.home),
  DrawerItem(title: 'Dashboard', icon: Icons.dashboard),
  DrawerItem(title: 'Laporan', icon: Icons.report),
];

final itemsSecond = [
  DrawerItem(title: 'Product', icon: Icons.archive),
  DrawerItem(title: 'Site', icon: Icons.factory),
  DrawerItem(title: 'PO & DO Page', icon: Icons.file_upload),
  DrawerItem(title: 'Invoice', icon: Icons.download),
];
