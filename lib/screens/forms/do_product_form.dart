import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

// import appbar function
import 'package:tric_duta_bangsa/widgets/app_bar/app_bar.dart';
import 'package:tric_duta_bangsa/widgets/random/status_pop_up.dart';

class DoProductForm extends StatefulWidget {
  final String doNumber;
  final String doId;
  const DoProductForm({super.key, required this.doNumber, required this.doId});

  @override
  State<DoProductForm> createState() => _DoProductFormState();
}

class _DoProductFormState extends State<DoProductForm> {
  final ApiServices _services = ApiServices();

  final TextEditingController _productNumberController =
      TextEditingController();
  final TextEditingController _productBrandController = TextEditingController();
  final TextEditingController _qtyController = TextEditingController(text: '0');

  List<DropdownMenuEntry<String>> productEntries = [];
  List<Map<String, dynamic>> _allData = [];

  int quantity = 0;
  // ignore: non_constant_identifier_names
  String UoM = '';
  String productId = '';

  @override
  void initState() {
    super.initState();
    _loadProducts(); // Fetch data when screen starts
  }

  // Fetch products from Database
  Future<void> _loadProducts() async {
    try {
      // Assuming your ProductService has a method like getAllProducts()
      // that returns a List of Maps or Models
      final List<Map<String, dynamic>> data = await _services
          .getAllSearchedData("product_limit_services.php", widget.doId);

      setState(() {
        _allData = data; // Store all data for later use

        productEntries = data.map((product) {
          return DropdownMenuEntry<String>(
            value: product['product_id'].toString(), // Use ID as value
            label: product['part_number'].toString(), // Use Name as label
          );
        }).toList();
      });
    } catch (e) {
      PopUp(status: 'error', title: 'Gagal Memuat Produk');
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.background,
      appBar: appBarCustom("PO:   ${widget.doNumber}", false),
      body: SingleChildScrollView(
        child: Container(
          margin: EdgeInsets.only(top: 16, left: 16, right: 16),
          color: AppTheme.box,
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Column(
              children: [
                // Site Information Section
                Row(
                  children: [
                    Text('Product       : ', style: TextStyle(fontSize: 24)),
                    DropdownMenu(
                      label: const Text('Nama Product'),
                      enableFilter: true,
                      enableSearch: true,
                      onSelected: (value) {
                        setState(() {
                          final selectedProduct = _allData.firstWhere(
                            (product) =>
                                product['product_id'].toString() == value,
                            orElse: () => {},
                          );

                          _productNumberController.text =
                              selectedProduct['part_number'].toString();
                          _productBrandController.text =
                              selectedProduct['brand'].toString();
                          UoM = selectedProduct['unit'].toString();
                          productId = selectedProduct['product_id'].toString();
                        });
                      },
                      dropdownMenuEntries: productEntries,
                    ),
                  ],
                ),
                SizedBox(height: 16),
                Column(
                  children: [
                    Text(
                      'Informasi Product',
                      style: TextStyle(
                        fontSize: 32,
                        fontWeight: FontWeight.normal,
                      ),
                    ),
                    TextField(
                      controller: _productNumberController,
                      enabled: false,
                      decoration: InputDecoration(
                        label: Text(
                          'Part Number',
                          style: TextStyle(fontSize: 24),
                        ),
                      ),
                    ),
                    SizedBox(height: 32),
                    TextField(
                      controller: _productBrandController,
                      enabled: false,
                      decoration: InputDecoration(
                        label: Text(
                          'Brand Product',
                          style: TextStyle(fontSize: 24),
                        ),
                      ),
                    ),
                  ],
                ),

                SizedBox(height: 32),
                //Product Stock Section
                Center(
                  child: Text(
                    'Informasi Stok Produk',
                    style: TextStyle(fontSize: 32),
                  ),
                ),
                SizedBox(height: 32),
                Center(child: Text('Quantity', style: TextStyle(fontSize: 24))),
                Row(
                  children: [
                    IconButton(
                      onPressed: () {
                        setState(() {
                          if (quantity > 0) {
                            quantity--;
                            _qtyController.text = quantity.toString();
                          }
                        });
                      },
                      icon: Icon(Icons.remove),
                    ),
                    Expanded(
                      child: TextField(
                        controller: _qtyController,
                        textAlign: TextAlign.center,
                        keyboardType: TextInputType.number,
                        inputFormatters: <TextInputFormatter>[
                          FilteringTextInputFormatter.digitsOnly,
                        ],
                        decoration: const InputDecoration(
                          hintText: 'Masukkan kuantitas produk',
                        ),
                        onChanged: (value) {
                          if (value.isEmpty) {
                            setState(() {
                              quantity = 0;
                            });
                          } else {
                            setState(() {
                              quantity = int.parse(value);
                            });
                          }
                        },
                      ),
                    ),
                    IconButton(
                      onPressed: () {
                        setState(() {
                          quantity++;
                          _qtyController.text = quantity.toString();
                        });
                      },
                      icon: Icon(Icons.add),
                    ),
                  ],
                ),
                SizedBox(height: 16),
                Text('Unit of Measurement : $UoM'),

                SizedBox(height: 32),
                Center(
                  child: ElevatedButton(
                    onPressed: () async {
                      Map<String, dynamic> data = {
                        "do_id": widget.doId,
                        "productId": productId,
                        "quantity": quantity,
                      };

                      _services.postData("do_item_services.php", data);
                    },
                    style: ButtonStyle(),
                    child: const Text('Submit'),
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
