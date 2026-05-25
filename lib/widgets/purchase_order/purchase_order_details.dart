import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'dart:async';

import 'package:tric_duta_bangsa/screens/forms/po_product_form.dart';
import 'package:tric_duta_bangsa/widgets/delivery_order/delivery_order_details.dart';
import 'package:tric_duta_bangsa/widgets/random/button.dart';
import 'package:tric_duta_bangsa/widgets/random/pop_up_form.dart';

class PurchaseOrderDetail extends StatefulWidget {
  final String poNumber;
  final bool makeDo;

  const PurchaseOrderDetail({
    super.key,
    required this.poNumber,
    this.makeDo = false,
  });

  @override
  State<PurchaseOrderDetail> createState() => _PurchaseOrderDetailState();
}

class _PurchaseOrderDetailState extends State<PurchaseOrderDetail> {
  final ApiServices _services = ApiServices();
  final Button _button = Button();

  List<dynamic> products = [];
  List<Map<String, dynamic>> po = [];
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
    _fetchpo(widget.poNumber);
    _fetchProduct(widget.poNumber);

    if (widget.makeDo) {
      _makeDo();
    }
  }

  Future<void> _fetchProduct([String? po]) async {
    setState(() => isLoading = true);

    List<dynamic> results;
    results = await _services.getAllSearchedData(
      "po_item_services.php",
      widget.poNumber,
    );

    setState(() {
      products = results;
      isLoading = false;
    });
  }

  // Fungsi utama ambil data & pencarian
  Future<void> _fetchpo([String? poNo]) async {
    setState(() => isLoading = true);

    List<Map<String, dynamic>> results;
    results = await _services.getAllSearchedData(
      "purchase_order_services.php",
      poNo ?? "",
    );

    setState(() {
      po = results;
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
          "${widget.poNumber}",
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

                const SizedBox(width: 20),

                _button.buildButton(
                  onPressed: () async {
                    FormPopUp.show(
                      context,
                      title: 'Add Product to PO ${widget.poNumber}',
                      child: PoProductForm(poNumber: widget.poNumber),
                    ).then((_) => _fetchProduct(widget.poNumber));
                  },
                  text: 'Add Button',
                  icon: Icons.add_box_outlined,
                ),

                const SizedBox(width: 20),

                _button.buildButton(
                  onPressed: () {
                    _makeDo();
                  },
                  text: 'Buat Delivery Order',
                  icon: Icons.upload_file,
                  context: context,
                ),
              ],
            ),
            const Divider(thickness: 2),
            // Table View for Products
            Expanded(
              child: Container(
                padding: EdgeInsets.all(8),
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

  Future<void> _makeDo() async {
    final response = await _services.postData("delivery_order_services.php", {
      "po_number": widget.poNumber,
    });

    if (response['status'] == 'success') {
      final doNumber = await response['do_number'];
      final doId = await response['id'];

      if (!mounted) return;
      Navigator.pushReplacement(
        context,
        MaterialPageRoute(
          builder: (context) => DeliveryOrderDetails(
            doNumber: doNumber.toString(),
            doId: doId.toString(),
          ),
        ),
      );
    }
  }

  Widget _buildHeaderInfo() {
    if (po.isEmpty) {
      return const Center(child: CircularProgressIndicator());
    }

    final data = po[0];

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
          _infoRow("Perusahaan", data['company_name'] ?? 'N/A'),
          _infoRow("Site", data['site_name'] ?? 'N/A'),
          _infoRow("Site Location:", data['site_address'] ?? 'N/A'),
          _infoRow("Date:", data['po_date'] ?? 'N/A'),
          _infoRow("Status:", "-"),
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
                      Text('Price Unit: ${item['unit_price']}'),
                      Text('Qty: ${item['quantity']} ${item['unit']}'),
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
                      SizedBox(width: 8),
                      _button.functionalButton(
                        onPressedD: () async {
                          await _services.deleteData(
                            "po_item_services.php",
                            item['po_item_id'].toString(),
                          );
                          _fetchProduct(); // Refresh data setelah delete
                        },
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
