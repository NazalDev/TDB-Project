import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

//function import
import 'package:tric_duta_bangsa/widgets/sidebar/navigation_provider.dart';
import 'package:tric_duta_bangsa/widgets/sidebar/drawer_items.dart';
import 'package:tric_duta_bangsa/widgets/sidebar/drawer_item.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

//import pages
import 'package:tric_duta_bangsa/screens/home_screen.dart';
import 'package:tric_duta_bangsa/screens/order_screens.dart';
import 'package:tric_duta_bangsa/screens/product_screen.dart';
import 'package:tric_duta_bangsa/screens/invoice_screen.dart';
import 'package:tric_duta_bangsa/screens/site_screen.dart';

class NavigationDrawerWidget extends StatelessWidget {
  final padding = EdgeInsets.symmetric(horizontal: 20);

  NavigationDrawerWidget({super.key});

  @override
  Widget build(BuildContext context) {
    final provider = Provider.of<NavigationProvider>(context);
    final isCollapsed = provider.isCollapse;

    return SizedBox(
      width: isCollapsed
          ? 100
          : null, //change the true one into Media Query for responsive widget
      child: Drawer(
        child: Container(
          color: AppTheme.greenMedium,
          child: Column(
            children: [
              Container(
                padding: EdgeInsets.symmetric(vertical: 16),
                width: double.infinity,
                color: AppTheme.greenDark,
                child: buildHeader(isCollapsed),
              ),
              buildList(items: itemsFirst, isCollapsed: isCollapsed, offset: 0),
              Divider(
                color: AppTheme.greenDark,
                thickness: 5,
                radius: BorderRadius.circular(8),
              ),
              buildList(
                items: itemsSecond,
                isCollapsed: isCollapsed,
                offset: itemsFirst.length,
              ),
              Spacer(),
              Divider(color: AppTheme.greenDark, thickness: 5),
              buildCollapseIcon(context, isCollapsed),
              const SizedBox(height: 12),
            ],
          ),
        ),
      ),
    );
  }

  Widget buildHeader(bool isCollapsed) => isCollapsed
      ? Icon(Icons.account_circle, size: 72, color: Colors.white)
      : Row(
          children: [
            const SizedBox(width: 24),
            Icon(Icons.account_circle, size: 72, color: Colors.white),
            const SizedBox(width: 16),
            Text(
              'TDB APP',
              style: TextStyle(
                fontSize: 24,
                color: Colors.white,
                fontWeight: FontWeight.bold,
              ),
            ),
          ],
        );

  Widget buildList({
    required bool isCollapsed,
    required List<DrawerItem> items,
    int offset = 0,
  }) => ListView.builder(
    shrinkWrap: true,
    primary: false,
    itemCount: items.length,
    itemBuilder: (context, index) {
      final item = items[index];

      return buildMenuItem(
        isCollapsed: isCollapsed,
        text: item.title,
        icon: item.icon,
        onClicked: () => selectItem(context, index + offset),
      );
    },
  );

  void selectItem(BuildContext context, int index) {
    Future navigateTo(page) => Navigator.of(
      context,
    ).push(MaterialPageRoute(builder: (context) => page));
    switch (index) {
      case 0:
        navigateTo(HomePage());
        break;
      case 3:
        navigateTo(ProductPage()); //change the page into the desired page
        break;
      case 4:
        navigateTo(SitePage());
      case 5:
        navigateTo(OrderPage());
        break;
      case 6:
        navigateTo(InvoicePage());
        break;
      // Add more cases for other indices if needed
    }
  }

  Widget buildMenuItem({
    required bool isCollapsed,
    required String text,
    required IconData icon,
    VoidCallback? onClicked,
  }) {
    final leading = Icon(icon, color: AppTheme.greenDark);

    return Material(
      color: Colors.transparent,
      child: Padding(
        padding: const EdgeInsets.all(8.0),
        child: isCollapsed
            ? Material(
                color: AppTheme.buttonNormal,
                shape: CircleBorder(),
                child: InkWell(
                  customBorder: CircleBorder(),
                  onTap: onClicked,
                  splashColor: AppTheme.clicked,
                  hoverColor: AppTheme.hover,

                  child: SizedBox(
                    width: 52,
                    height: 52,
                    child: Center(child: leading),
                  ),
                ),
              )
            : ListTile(
                hoverColor: AppTheme.hover,
                splashColor: AppTheme.clicked,
                onTap: onClicked,
                leading: leading,
                title: Text(
                  text,
                  style: TextStyle(
                    fontFamily: 'Montserrat',
                    fontWeight: FontWeight.bold,
                    color: AppTheme.textPrimary,
                  ),
                ),
                tileColor: AppTheme.buttonNormal,
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(16),
                ),
              ),
      ),
    );
  }

  Widget buildCollapseIcon(BuildContext context, bool isCollapsed) {
    final double size = 52;
    final icon = isCollapsed ? Icons.arrow_forward_ios : Icons.arrow_back_ios;
    final alignment = isCollapsed ? Alignment.center : Alignment.centerRight;
    final margin = isCollapsed ? null : EdgeInsets.only(right: 16);
    final width = isCollapsed ? double.infinity : size;

    return Container(
      alignment: alignment,
      margin: margin,
      child: Material(
        color: Colors.transparent,
        shape: CircleBorder(),
        child: InkWell(
          customBorder: CircleBorder(),
          splashColor: AppTheme.clicked,
          hoverColor: AppTheme.hover,
          onTap: () {
            final provider = Provider.of<NavigationProvider>(
              context,
              listen: false,
            );
            provider.toggleCollapsed();
          },
          child: SizedBox(
            width: width,
            height: size,
            child: Icon(icon, color: AppTheme.buttonNormal),
          ),
        ),
      ),
    );
  }
}
