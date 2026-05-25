import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'package:tric_duta_bangsa/widgets/app_bar/app_bar.dart';
import 'package:tric_duta_bangsa/widgets/random/button.dart';
import 'package:tric_duta_bangsa/widgets/random/pop_up_form.dart';
import 'package:tric_duta_bangsa/widgets/random/status_pop_up.dart';
import 'package:tric_duta_bangsa/widgets/sidebar/navigation_sidebar.dart';
import 'package:tric_duta_bangsa/screens/forms/product_form.dart';
import 'dart:async'; // Perlu untuk Timer (Debounce)

class ProductPage extends StatefulWidget {
  const ProductPage({super.key});

  @override
  State<ProductPage> createState() => _ProductPageState();
}

class _ProductPageState extends State<ProductPage> {
  final ApiServices _service = ApiServices();

  // Data State
  List<dynamic> products = [];
  List<bool> selectedItems = [];
  bool isSelectionMode = false;
  bool isLoading = false;
  String actionType = '';
  Timer? _debounce;

  @override
  void initState() {
    super.initState();
    _fetchProducts(); // Ambil data saat pertama kali buka
  }

  // Fungsi utama ambil data & pencarian
  Future<void> _fetchProducts([String? query]) async {
    setState(() => isLoading = true);

    List<dynamic> results;
    if (query == null || query.isEmpty) {
      results = await _service.getAllData("product_services.php");
    } else {
      results = await _service.getAllSearchedData(
        "product_services.php",
        query,
      );
    }

    setState(() {
      products = results;
      selectedItems = List.generate(results.length, (index) => false);
      isLoading = false;
    });
  }

  // Fungsi yang dipanggil oleh SearchBar
  void _onSearchChanged(String value) {
    if (_debounce?.isActive ?? false) _debounce!.cancel();
    _debounce = Timer(const Duration(milliseconds: 500), () {
      _fetchProducts(value);
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
      appBar: appBarCustom("Products", true),
      drawer: NavigationDrawerWidget(),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Container(
          decoration: BoxDecoration(
            color: AppTheme.box,
            borderRadius: BorderRadius.circular(8),
          ),
          child: Column(
            children: [
              // Header: Add & Search
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
                  padding: const EdgeInsets.all(8.0),
                  margin: const EdgeInsets.symmetric(horizontal: 8),
                  decoration: BoxDecoration(
                    border: Border.all(color: AppTheme.textPrimary, width: 3),
                    borderRadius: BorderRadius.circular(8),
                  ),
                  child: isLoading
                      ? const Center(child: CircularProgressIndicator())
                      : products.isEmpty
                      ? const Center(child: Text("Data tidak ditemukan"))
                      : buildProductList(),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildSearchBar() {
    return Container(
      decoration: BoxDecoration(
        border: Border.all(color: AppTheme.textPrimary, width: 3),
        borderRadius: BorderRadius.circular(8),
      ),
      child: SearchBar(
        onChanged: _onSearchChanged,
        leading: const Icon(Icons.search),
        hintText: 'Search product...',
        constraints: const BoxConstraints(
          minHeight: 40,
          maxHeight: 40, // Disamakan agar rapi
          maxWidth: 200,
        ),
        backgroundColor: WidgetStateProperty.all(AppTheme.buttonNormal),
        shadowColor: WidgetStateProperty.all(Colors.transparent),
        elevation: WidgetStateProperty.all(0),
        shape: WidgetStateProperty.all(
          RoundedRectangleBorder(borderRadius: BorderRadius.circular(5)),
        ),
      ),
    );
  }

  Widget buildProductList() {
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
                        "${item['product_name'] ?? 'N/A'} (${item['brand'] ?? 'N/A'})",
                      ),
                    ],
                  ),
                  Button().functionalButton(
                    onPressedV: () {},
                    onPressedE: () {},
                    onPressedD: () async {
                      final res = await _service.deleteData(
                        "product_services.php",
                        item['product_id'].toString(),
                      );

                      PopUp(status: res['status'], title: res['message']);

                      await _fetchProducts(); // Refresh data setelah delete
                    },
                  ),
                ],
              ),
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 16),
                height: 24,
                decoration: BoxDecoration(
                  color: AppTheme.headingFooterMainBox,
                  borderRadius: BorderRadius.all(Radius.circular(8)),
                ),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  crossAxisAlignment: CrossAxisAlignment.center,
                  children: [
                    Text(
                      'Stock : ${item['stock_quantity'] ?? 'N/A'} ${item['unit'] ?? ''}',
                    ),
                  ],
                ),
              ),
            ],
          ),
        );
      },
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
            title: "Add Product",
            child: const ProductForm(),
            onClose: _fetchProducts, // Refresh data setelah kembali dari form
          );
        },
        style: ElevatedButton.styleFrom(
          backgroundColor: AppTheme.buttonNormal,
          shape: const RoundedRectangleBorder(borderRadius: BorderRadius.zero),
        ),
        icon: const Icon(Icons.add_box_outlined, color: Colors.black),
        label: const Text(
          'ADD Product',
          style: TextStyle(color: Colors.black, fontWeight: FontWeight.bold),
        ),
      ),
    );
  }
}
