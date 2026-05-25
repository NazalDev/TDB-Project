import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/widgets/random/button.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'dart:async';

// import form
import 'package:tric_duta_bangsa/screens/forms/po_form.dart';

//import detail page
import 'package:tric_duta_bangsa/widgets/purchase_order/purchase_order_details.dart';
import 'package:tric_duta_bangsa/widgets/random/pop_up_form.dart';

class PurchaseOrderMain extends StatefulWidget {
  const PurchaseOrderMain({super.key});

  @override
  State<PurchaseOrderMain> createState() => _PurchaseOrderMainState();
}

class _PurchaseOrderMainState extends State<PurchaseOrderMain> {
  final ApiServices _services = ApiServices();
  final Button _button = Button();

  List<dynamic> po = [];
  bool isLoading = false;
  Timer? _debounce;

  @override
  void initState() {
    super.initState();
    _fetchpo();
  }

  // Fungsi utama ambil data & pencarian
  Future<void> _fetchpo([String? query]) async {
    setState(() => isLoading = true);

    List<dynamic> results;
    if (query == null || query.isEmpty) {
      results = await _services.getAllData("purchase_order_services.php");
    } else {
      results = await _services.getAllSearchedData(
        "purchase_order_services.php",
        query,
      );
    }

    setState(() {
      po = results;
      isLoading = false;
    });
  }

  // Fungsi yang dipanggil oleh SearchBar
  void _onSearchChanged(String value) {
    if (_debounce?.isActive ?? false) _debounce!.cancel();
    _debounce = Timer(const Duration(milliseconds: 500), () {
      _fetchpo(value);
    });
  }

  @override
  void dispose() {
    _debounce?.cancel();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Container(
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
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [_buildAddButton(), _buildSearchBar()],
            ),
          ),
          // List Content
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
                  : po.isEmpty
                  ? const Center(child: Text("Data tidak ditemukan"))
                  : _buildPoList(),
            ),
          ),
        ],
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

  Widget _buildAddButton() {
    return Container(
      decoration: BoxDecoration(
        border: Border.all(color: AppTheme.textPrimary, width: 3),
        borderRadius: BorderRadius.circular(8),
      ),
      child: ElevatedButton.icon(
        onPressed: () async {
          FormPopUp.show(
            context,
            title: 'Add Purchase Order',
            child: PurchaseOrderForm(),
            onClose: _fetchpo, // refreshes your list automatically
          ); // Refresh data setelah kembali dari form
        },
        style: ElevatedButton.styleFrom(
          backgroundColor: AppTheme.buttonNormal,
          shape: const RoundedRectangleBorder(borderRadius: BorderRadius.zero),
        ),
        icon: const Icon(Icons.add_box_outlined, color: Colors.black),
        label: const Text(
          'ADD PO',
          style: TextStyle(color: Colors.black, fontWeight: FontWeight.bold),
        ),
      ),
    );
  }

  Widget _buildPoList() {
    return ListView.separated(
      itemCount: po.length,
      separatorBuilder: (context, index) => const SizedBox(height: 8),
      itemBuilder: (context, index) {
        final item = po[index];
        return GestureDetector(
          onTap: () {
            Navigator.push(
              context,
              MaterialPageRoute(
                builder: (context) =>
                    PurchaseOrderDetail(poNumber: item['po_number']),
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
                          item['po_number'] ?? 'N/A',
                          style: TextStyle(
                            fontWeight: FontWeight.w500,
                            fontSize: 16,
                          ),
                        ),
                        Text('Tanggal masuk: ${item['po_date']}'),
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
                            builder: (context) => PurchaseOrderDetail(
                              poNumber: item['po_number'],
                            ),
                          ),
                        );
                      },
                      onPressedE: () {},
                      onPressedD: () async {
                        await _services.deleteData(
                          "purchase_order_services.php",
                          item['po_number'].toString(),
                        );
                        await _fetchpo(); // Refresh data setelah delete
                      },
                      onPressedM: () {
                        Navigator.push(
                          context,
                          MaterialPageRoute(
                            builder: (context) => PurchaseOrderDetail(
                              poNumber: item['po_number'],
                              makeDo: true,
                            ),
                          ),
                        );
                      },
                      mlabel: 'Generate DO',
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
