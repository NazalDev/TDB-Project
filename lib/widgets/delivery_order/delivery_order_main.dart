import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/widgets/random/button.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'dart:async';

// import detail page
import 'package:tric_duta_bangsa/widgets/delivery_order/delivery_order_details.dart';
import 'package:tric_duta_bangsa/widgets/random/status_pop_up.dart';

class DeliveryOrderMain extends StatefulWidget {
  const DeliveryOrderMain({super.key});

  @override
  State<DeliveryOrderMain> createState() => _DeliveryOrderMainState();
}

class _DeliveryOrderMainState extends State<DeliveryOrderMain> {
  final ApiServices _services = ApiServices();
  final Button _button = Button();

  List<dynamic> deliveryOrders = [];
  bool isLoading = false;
  Timer? _debounce;

  @override
  void initState() {
    super.initState();
    _fetchdo();
  }

  // Fungsi utama ambil data & pencarian
  Future<void> _fetchdo([String? query]) async {
    setState(() => isLoading = true);

    List<dynamic> results;
    if (query == null || query.isEmpty) {
      results = await _services.getAllData("delivery_order_services.php");
    } else {
      results = await _services.getAllSearchedData(
        "delivery_order_services.php",
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
      _fetchdo(value);
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
      color: AppTheme.box,
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
                  : _buildDoList(),
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
        hintText: 'Search Delivery Order...',
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

  Widget _buildDoList() {
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
                builder: (context) => DeliveryOrderDetails(
                  doNumber: item['do_number'],
                  doId: item['do_id'],
                ),
              ),
            );
          },
          child: Container(
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
                          item['do_number'] ?? 'No Name',
                          style: TextStyle(
                            fontWeight: FontWeight.w500,
                            fontSize: 16,
                          ),
                        ),
                        Text('Tanggal dibuat: ${item['do_date']}'),
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
                            builder: (context) => DeliveryOrderDetails(
                              doNumber: item['do_number'],
                              doId: item['do_id'],
                            ),
                          ),
                        );
                      },
                      onPressedD: () async {
                        await _services.deleteData(
                          "delivery_order_services.php",
                          item['do_id'].toString(),
                        );
                        await _fetchdo(); // Refresh data setelah delete
                      },

                      onPressedM: () {
                        item['po_is_deleted'] == 1 ||
                                item['po_is_deleted'] == '1' ||
                                item['po_is_deleted'] == true
                            ? showDialog(
                                context: context,
                                builder: (context) => PopUp(
                                  status: 'error',
                                  title: 'Gagal membuat Invoice',
                                  description: 'Purchase Order telah dihapus',
                                ),
                              )
                            : Navigator.push(
                                context,
                                MaterialPageRoute(
                                  builder: (context) => DeliveryOrderDetails(
                                    doNumber: item['do_number'],
                                    doId: item['do_id'],
                                    openForInvoice: true,
                                  ),
                                ),
                              );
                      },
                      mlabel: 'Generate Invoice',
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
