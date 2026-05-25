import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'package:tric_duta_bangsa/widgets/random/button.dart';
import 'dart:async';

//function import
import 'package:tric_duta_bangsa/widgets/sidebar/navigation_sidebar.dart';
import 'package:tric_duta_bangsa/widgets/app_bar/app_bar.dart';

import 'package:tric_duta_bangsa/widgets/invoice/invoice_details.dart';

class InvoicePage extends StatefulWidget {
  const InvoicePage({super.key});

  @override
  State<InvoicePage> createState() => _InvoicePageState();
}

class _InvoicePageState extends State<InvoicePage> {
  final ApiServices _services = ApiServices();
  final Button _button = Button();

  List<dynamic> deliveryOrders = [];
  bool isLoading = false;
  Timer? _debounce;

  @override
  void initState() {
    super.initState();
    _fetchInvoice();
  }

  // Fungsi utama ambil data & pencarian
  Future<void> _fetchInvoice([String? query]) async {
    setState(() => isLoading = true);

    List<dynamic> results;
    if (query == null || query.isEmpty) {
      results = await _services.getAllData("invoice_services.php");
    } else {
      results = await _services.getAllSearchedData(
        "invoice_services.php",
        query,
      );
    }

    setState(() {
      deliveryOrders = results;
      isLoading = false;
    });
  }

  // Fungsi yang dipanggil oleh SearchBar
  void _onSearchChanged(String value) {
    if (_debounce?.isActive ?? false) _debounce!.cancel();
    _debounce = Timer(const Duration(milliseconds: 500), () {
      _fetchInvoice(value);
    });
  }

  @override
  void dispose() {
    _debounce?.cancel();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.background,
      appBar: appBarCustom("INVOICE", true),
      drawer: NavigationDrawerWidget(),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Container(
          decoration: BoxDecoration(
            color: AppTheme.box,
            borderRadius: BorderRadius.circular(8),
          ),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              Padding(
                padding: const EdgeInsets.all(8.0),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.end,
                  children: [_buildSearchBar()],
                ),
              ),
              Expanded(
                child: Container(
                  padding: const EdgeInsets.all(8),
                  margin: const EdgeInsets.symmetric(horizontal: 8),
                  decoration: BoxDecoration(
                    border: Border.all(color: AppTheme.textPrimary, width: 3),
                    borderRadius: BorderRadius.circular(8),
                  ),
                  child: isLoading
                      ? const Center(child: CircularProgressIndicator())
                      : deliveryOrders.isEmpty
                      ? const Center(child: Text("Data tidak ditemukan"))
                      : _buildInvoiceList(),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Container _buildSearchBar() {
    return Container(
      decoration: BoxDecoration(
        border: Border.all(color: AppTheme.textPrimary, width: 3),
        borderRadius: BorderRadius.circular(8),
      ),
      child: SearchBar(
        onChanged: _onSearchChanged,
        leading: const Icon(Icons.search),
        hintText: 'Search Po...',
        constraints: const BoxConstraints(
          minHeight: 40,
          minWidth: 25,
          maxHeight: 250,
          maxWidth: 250,
        ),
        backgroundColor: WidgetStateProperty.all(AppTheme.buttonNormal),
        shadowColor: WidgetStateProperty.all(Colors.transparent),
        shape: WidgetStateProperty.all(
          RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
        ),
      ),
    );
  }

  Widget _buildInvoiceList() {
    return ListView.separated(
      itemCount: deliveryOrders.length,
      separatorBuilder: (context, index) => const SizedBox(height: 8),
      itemBuilder: (context, index) {
        final item = deliveryOrders[index];
        return GestureDetector(
          onTap: () {
            Navigator.push(
              context,
              MaterialPageRoute(
                builder: (context) => InvoiceDetails(
                  invoiceId: item['invoice_id'],
                  invoiceNumber: item['invoice_number'],
                ),
              ),
            );
          },
          child: Container(
            padding: const EdgeInsets.all(12),
            decoration: BoxDecoration(
              border: BoxBorder.all(color: AppTheme.greenMedium, width: 2),
              borderRadius: BorderRadius.all(Radius.circular(8)),
            ),

            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              spacing: 2,
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          item['invoice_number'] ?? 'N/A',
                          style: TextStyle(
                            fontWeight: FontWeight.w500,
                            fontSize: 16,
                          ),
                        ),
                        Text('Term of Payment: ${item['due_date']}'),
                        Text('Po: ${item['po_number']}'),
                        Text(
                          'Site: ${item['company_short_name']} ${item['site_name']}',
                        ),
                      ],
                    ),
                    _button.functionalButton(
                      onPressedV: () {
                        Navigator.push(
                          context,
                          MaterialPageRoute(
                            builder: (context) => InvoiceDetails(
                              invoiceId: item['invoice_id'],
                              invoiceNumber: item['invoice_number'],
                            ),
                          ),
                        );
                      },
                      onPressedD: () async {
                        await _services.deleteData(
                          "invoice_services.php",
                          item['invoice_number'],
                        );
                        await _fetchInvoice();
                      },
                    ),
                  ],
                ),
              ],
            ),
          ),
        );
      },
    );
  }
}
