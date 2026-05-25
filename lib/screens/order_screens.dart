import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

//function import
import 'package:tric_duta_bangsa/widgets/app_bar/app_bar.dart';
import 'package:tric_duta_bangsa/widgets/sidebar/navigation_sidebar.dart';

//widget import
import 'package:tric_duta_bangsa/widgets/purchase_order/purchase_order_main.dart';
import 'package:tric_duta_bangsa/widgets/delivery_order/delivery_order_main.dart';

class OrderPage extends StatelessWidget {
  const OrderPage({super.key});

  @override
  Widget build(BuildContext context) {
    return DefaultTabController(
      length: 2,
      child: Scaffold(
        backgroundColor: AppTheme.background,
        appBar: appBarCustom("Purchase & Delivery Order", true),
        drawer: NavigationDrawerWidget(),
        body: Padding(
          padding: const EdgeInsets.all(16.0),
          child: Column(
            children: [
              Container(
                decoration: BoxDecoration(
                  color: AppTheme.headingFooterMainBox,
                  borderRadius: BorderRadius.circular(8),
                ),
                child: SizedBox(
                  height: 76,
                  child: TabBar(
                    dividerColor: AppTheme.clicked,
                    dividerHeight: 3,
                    indicatorColor: AppTheme.greenMedium,
                    indicatorWeight: 5,
                    indicatorPadding: EdgeInsetsGeometry.only(bottom: 2),
                    tabs: [
                      Tab(
                        child: Row(
                          crossAxisAlignment: CrossAxisAlignment.center,
                          mainAxisAlignment: MainAxisAlignment.center,
                          children: [
                            Padding(
                              padding: const EdgeInsets.only(right: 8),
                              child: Icon(
                                Icons.file_download,
                                color: AppTheme.greenMedium,
                              ),
                            ),
                            Padding(
                              padding: const EdgeInsets.only(left: 8),
                              child: Text(
                                'Purchase Order',
                                style: TextStyle(
                                  color: AppTheme.textPrimary,
                                  fontSize: 16,
                                  fontWeight: FontWeight.bold,
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                      Tab(
                        child: Row(
                          crossAxisAlignment: CrossAxisAlignment.center,
                          mainAxisAlignment: MainAxisAlignment.center,
                          children: [
                            Padding(
                              padding: const EdgeInsets.only(right: 8),
                              child: Icon(
                                Icons.file_upload,
                                color: AppTheme.greenMedium,
                              ),
                            ),
                            Padding(
                              padding: const EdgeInsets.only(left: 8),
                              child: Text(
                                'Delivery Order',
                                style: TextStyle(
                                  color: AppTheme.textPrimary,
                                  fontSize: 16,
                                  fontWeight: FontWeight.bold,
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                    ],
                  ),
                ),
              ),
              Expanded(
                child: TabBarView(
                  children: [PurchaseOrderMain(), DeliveryOrderMain()],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
