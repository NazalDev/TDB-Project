import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'package:tric_duta_bangsa/widgets/random/button.dart';
import 'package:tric_duta_bangsa/widgets/random/pop_up_form.dart';
import 'dart:async';

//function import
import 'package:tric_duta_bangsa/widgets/sidebar/navigation_sidebar.dart';
import 'package:tric_duta_bangsa/widgets/app_bar/app_bar.dart';

//import form
import 'package:tric_duta_bangsa/screens/forms/site_form.dart';

class SitePage extends StatefulWidget {
  const SitePage({super.key});

  @override
  State<SitePage> createState() => _SitePageState();
}

class _SitePageState extends State<SitePage> {
  final ApiServices _services = ApiServices();
  final Button _button = Button();

  // Data State
  List<dynamic> sites = [];
  bool isLoading = false;
  Timer? _debounce;

  @override
  void initState() {
    super.initState();
    _fetchsites();
  }

  // Fungsi utama ambil data & pencarian
  Future<void> _fetchsites([String? query]) async {
    setState(() => isLoading = true);

    List<dynamic> results;
    if (query == null || query.isEmpty) {
      results = await _services.getAllData("site_services.php");
    } else {
      results = await _services.getAllSearchedData("site_services.php", query);
    }

    setState(() {
      sites = results;
      isLoading = false;
    });
  }

  // Fungsi yang dipanggil oleh SearchBar
  void _onSearchChanged(String value) {
    if (_debounce?.isActive ?? false) _debounce!.cancel();
    _debounce = Timer(const Duration(milliseconds: 500), () {
      _fetchsites(value);
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
      appBar: appBarCustom("SITE", true),
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
                      : sites.isEmpty
                      ? const Center(child: Text("Data tidak ditemukan"))
                      : _buildProductList(),
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
        hintText: 'Search Site...',
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
            title: 'Add Site',
            child: SiteForm(),
            onClose: _fetchsites, // refreshes your list automatically
          ); // Refresh data setelah kembali dari form // Refresh data setelah kembali dari form
        },
        style: ElevatedButton.styleFrom(
          backgroundColor: AppTheme.buttonNormal,
          shape: const RoundedRectangleBorder(borderRadius: BorderRadius.zero),
        ),
        icon: const Icon(Icons.add_box_outlined, color: Colors.black),
        label: const Text(
          'ADD SITE',
          style: TextStyle(color: Colors.black, fontWeight: FontWeight.bold),
        ),
      ),
    );
  }

  Widget _buildProductList() {
    return ListView.separated(
      itemCount: sites.length,
      separatorBuilder: (context, index) => const SizedBox(height: 8),
      itemBuilder: (context, index) {
        final item = sites[index];
        return Container(
          padding: const EdgeInsets.all(12),
          decoration: BoxDecoration(
            color:
                item['po_is_deleted'] == 1 ||
                    item['po_is_deleted'] == '1' ||
                    item['po_is_deleted'] == true
                ? AppTheme.accentRed
                : AppTheme.box,
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
                        item['company_name'] ?? 'No Name',
                        style: TextStyle(
                          fontWeight: FontWeight.w500,
                          fontSize: 16,
                        ),
                      ),
                      Text('Site: ${item['site_name']}'),
                      Text('Alamat Site: ${item['site_address']}'),
                    ],
                  ),
                  _button.functionalButton(
                    onPressedD: () {},
                    onPressedE: () {},
                    onPressedV: () {},
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
