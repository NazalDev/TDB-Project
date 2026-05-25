import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'dart:async';

class InvoiceDetails extends StatefulWidget {
  final String invoiceNumber;
  final String invoiceId;
  const InvoiceDetails({
    super.key,
    required this.invoiceId,
    required this.invoiceNumber,
  });

  @override
  State<InvoiceDetails> createState() => _InvoiceDetailsState();
}

class _InvoiceDetailsState extends State<InvoiceDetails> {
  final ApiServices _services = ApiServices();

  bool isSelectionMode = false;
  String actionType = '';
  List<dynamic> products = [];
  List<dynamic> invoice = [];
  List<bool> selectedItems = [];
  bool isLoading = false;

  final TextEditingController _notesController = TextEditingController();

  @override
  void dispose() {
    _notesController.dispose(); // Always clean up controllers
    super.dispose();
  }

  @override
  void initState() {
    super.initState();
    _initData();
  }

  Future<void> _initData() async {
    setState(() => isLoading = true);

    List<dynamic> results;
    results = await _services.getAllSearchedData(
      "invoice_item_services.php",
      widget.invoiceId,
    );

    List<dynamic> results2;
    results2 = await _services.getAllSearchedData(
      "invoice_services.php",
      widget.invoiceId,
    );

    setState(() {
      products = results;
      invoice = results2;
      isLoading = false;
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.box,
      appBar: AppBar(
        title: Text(
          // ignore: unnecessary_string_interpolations
          "${widget.invoiceNumber}",
          style: TextStyle(color: AppTheme.textPrimary),
        ),
        backgroundColor: AppTheme.headingFooterMainBox,
        iconTheme: IconThemeData(color: AppTheme.textPrimary),
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            // Header Information Section
            _buildHeaderInfo(),
            const SizedBox(height: 20),
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                Text(
                  "Product List",
                  style: TextStyle(
                    fontSize: 20,
                    fontWeight: FontWeight.bold,
                    color: AppTheme.textPrimary,
                  ),
                ),
              ],
            ),
            const Divider(thickness: 2),
            // Table View for Products
            Expanded(
              child: Container(
                padding: const EdgeInsets.all(8),
                decoration: BoxDecoration(
                  border: Border.all(color: AppTheme.textPrimary, width: 2),
                  borderRadius: BorderRadius.circular(8),
                ),
                child: isLoading
                    ? const Center(child: CircularProgressIndicator())
                    : products.isEmpty
                    ? const Center(child: Text("Data tidak ditemukan"))
                    : _buildProductList(),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildHeaderInfo() {
    if (invoice.isEmpty) {
      return const Center(child: CircularProgressIndicator());
    }

    final data = invoice[0];

    return Container(
      padding: EdgeInsets.all(12),
      decoration: BoxDecoration(
        // ignore: deprecated_member_use
        color: AppTheme.buttonNormal.withOpacity(0.5),
        borderRadius: BorderRadius.circular(8),
        border: Border.all(color: AppTheme.textPrimary),
      ),
      child: Column(
        children: [
          _infoRow("Purchase Order:", data['po_number'] ?? 'N/A'),
          _infoRow("Perusahaan:", data['company_name'] ?? 'N/A'),
          _infoRow("Site:", data['site_name'] ?? 'N/A'),
          _infoRow("Site Location:", data['site_address'] ?? 'N/A'),
          _infoRow("Date:", data['invoice_date'] ?? 'N/A'),
          _infoRow("Term of Payment:", data['due_date'] ?? 'N/A'),
        ],
      ),
    );
  }

  Widget _infoRow(String label, String value) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4.0),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(
            label,
            style: TextStyle(
              fontWeight: FontWeight.bold,
              color: AppTheme.textPrimary,
            ),
          ),
          Text(value, style: TextStyle(color: AppTheme.textPrimary)),
        ],
      ),
    );
  }

  Widget _buildProductList() {
    return ListView.separated(
      itemCount: products.length,
      separatorBuilder: (context, index) => const SizedBox(height: 8),
      itemBuilder: (context, index) {
        final item = products[index];
        return Container(
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
                        item['part_number'] ?? 'N/A',
                        style: TextStyle(
                          fontWeight: FontWeight.w500,
                          fontSize: 16,
                        ),
                      ),
                      Text(
                        'Barang yang dikirim: ${item['quantity']} ${item['unit']}',
                      ),
                      Text('Unit price: ${item['price']}'),
                    ],
                  ),
                  Row(
                    children: [
                      Text(
                        "Rp ${item['total']}",
                        style: TextStyle(
                          color: AppTheme.textPrimary,
                          fontWeight: FontWeight.bold,
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ],
          ),
        );
      },
    );
  }
}
