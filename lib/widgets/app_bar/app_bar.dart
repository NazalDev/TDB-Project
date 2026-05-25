import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

// Header
AppBar appBarCustom(String title, bool iconProfile) {
  return AppBar(
    //automaticallyImplyLeading: false,
    toolbarHeight: 80,
    title: Text(
      title.toUpperCase(),
      style: TextStyle(
        color: AppTheme.textPrimary,
        fontSize: 32,
        fontWeight: FontWeight.bold,
      ),
    ),
    backgroundColor: AppTheme.header,
    centerTitle: true,
    actions: iconProfile
        ? [
            Container(
              margin: EdgeInsets.all(10),
              width: 48,
              height: 48,
              decoration: BoxDecoration(
                color: Colors.white,
                borderRadius: BorderRadius.circular(50),
              ),
              child: _actionButton(
                () => null,
                Icons.person,
                AppTheme.greenMedium,
                32,
              ),
            ),
          ]
        : null,
  );
}

//Button with Icon Widget
IconButton _actionButton(
  Function()? onPressed,
  IconData icon,
  Color color,
  double size,
) {
  return IconButton(
    onPressed: onPressed,
    iconSize: size,
    highlightColor: Colors.transparent,
    icon: Icon(icon, color: color),
  );
}

/*
Container(
        margin: EdgeInsets.all(10),
        width: 48,
        height: 48,
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.circular(50),
        ),
        child: _actionButton(
          () => null,
          Icons.person,
          AppTheme.greenMedium,
          32,
        ),
      ),
    */
