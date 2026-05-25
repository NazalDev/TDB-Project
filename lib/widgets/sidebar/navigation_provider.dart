import 'package:flutter/material.dart';

class NavigationProvider extends ChangeNotifier {
  bool _isCollapsed = false;

  bool get isCollapse => _isCollapsed;

  void toggleCollapsed() {
    _isCollapsed = !_isCollapsed;
    notifyListeners();
  }
}
