import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

//import sql services
import 'package:tric_duta_bangsa/widgets/random/status_pop_up.dart';

class ProductForm extends StatefulWidget {
  final Map<String, dynamic>? product;

  const ProductForm({super.key, this.product});

  @override
  State<ProductForm> createState() => _ProductFormState();
}

class _ProductFormState extends State<ProductForm> {
  // variables
  int quantity = 0;
  bool _isNotLocked = true;
  int productId = 0;

  String selectedUom = 'Pcs'; // Nilai default
  final List<String> uomOptions = ['Pcs', 'Box', 'Meter'];

  //services
  final ApiServices _service = ApiServices();

  // load data from sql
  List<DropdownMenuEntry<String>> productEntries = [];
  List<Map<String, dynamic>> _allData = [];
  bool isLoading = true;

  //controller
  final TextEditingController _qtyController = TextEditingController(text: '0');
  final TextEditingController _partNumController = TextEditingController();
  final TextEditingController _nameController = TextEditingController();
  final TextEditingController _brandController = TextEditingController();
  final TextEditingController _remarksController = TextEditingController();

  @override
  void initState() {
    super.initState();
    if (widget.product != null) {
      productId =
          int.tryParse(widget.product!['product_id']?.toString() ?? '0') ?? 0;
      _partNumController.text =
          widget.product!['part_number']?.toString() ?? '';
      _nameController.text = widget.product!['product_name']?.toString() ?? '';
      _brandController.text = widget.product!['brand']?.toString() ?? '';
      _remarksController.text = widget.product!['remarks']?.toString() ?? '';

      // Set Quantity and UoM
      quantity =
          int.tryParse(widget.product!['stock_quantity']?.toString() ?? '0') ??
          0;
      _qtyController.text = quantity.toString();
      selectedUom = widget.product!['unit']?.toString() ?? 'Pcs';
    }
    _loadProducts(); // Fetch data when screen starts
  }

  // Fetch products from Database
  Future<void> _loadProducts() async {
    try {
      // Assuming your ProductService has a method like getAllProducts()
      // that returns a List of Maps or Models
      final List<Map<String, dynamic>> data = await _service.getAllData(
        'product_services.php',
      );

      setState(() {
        _allData = data; // Store all data for later use

        productEntries = data.map((product) {
          return DropdownMenuEntry<String>(
            value: product['part_number'].toString(), // Use ID as value
            label: product['part_number'].toString(), // Use Name as label
          );
        }).toList();

        // Add the "Add New Product" option at the end or beginning
        productEntries.add(
          const DropdownMenuEntry(
            value: 'openForm',
            label: 'Tambah Produk Baru',
          ),
        );

        isLoading = false;
      });
    } catch (e) {
      if (!mounted) return;
      showDialog(
        context: context,
        builder: (context) => PopUp(status: 'error', title: 'Gagal Load Data'),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    return Container(
      margin: EdgeInsets.only(top: 16, left: 16, right: 16),
      color: AppTheme.box,
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: ListView(
          physics: BouncingScrollPhysics(),
          shrinkWrap: true,
          children: [
            if (widget.product == null)
              // Product Preset Section
              Row(
                children: [
                  Text('Produk : ', style: TextStyle(fontSize: 24)),
                  DropdownMenu(
                    label: const Text('Nama Produk'),
                    enableFilter: true,
                    enableSearch: true,
                    onSelected: (value) {
                      setState(() {
                        if (value == 'openForm') {
                          _partNumController.clear();
                          _nameController.clear();
                          _brandController.clear();
                          _remarksController.clear();
                          _isNotLocked = true;
                        } else if (value != null && value != 'openForm') {
                          final selectedProduct = _allData.firstWhere(
                            (product) =>
                                product['part_number'].toString() == value,
                            orElse: () => {},
                          );
                          _partNumController.text =
                              selectedProduct['part_number'].toString();
                          _nameController.text = selectedProduct['product_name']
                              .toString();
                          _brandController.text = selectedProduct['brand']
                              .toString();
                          selectedUom = selectedProduct['unit'].toString();
                          _isNotLocked = false;
                        }
                      });
                    },
                    dropdownMenuEntries: productEntries,
                  ),
                ],
              ),
            SizedBox(height: 32),

            // Informasi Product Section
            Center(
              child: Text(
                'Informasi Produk',
                style: TextStyle(fontSize: 32, fontWeight: FontWeight.normal),
              ),
            ),
            TextField(
              controller: _partNumController,
              keyboardType: TextInputType.text,
              enabled: _isNotLocked,
              decoration: InputDecoration(
                label: Text('Part number', style: TextStyle(fontSize: 24)),
              ),
            ),
            SizedBox(height: 32),
            TextField(
              controller: _nameController,
              keyboardType: TextInputType.text,
              enabled: _isNotLocked,
              decoration: InputDecoration(
                label: Text(
                  'Deskripsi / Nama Part',
                  style: TextStyle(fontSize: 24),
                ),
              ),
            ),
            SizedBox(height: 32),
            TextField(
              controller: _brandController,
              keyboardType: TextInputType.text,
              enabled: _isNotLocked,
              decoration: InputDecoration(
                label: Text('Brand / Merk', style: TextStyle(fontSize: 24)),
              ),
            ),

            SizedBox(height: 32),

            //Product Information Section
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
            SizedBox(height: 32),
            Row(
              crossAxisAlignment: CrossAxisAlignment.end,
              children: [
                Padding(
                  padding: const EdgeInsets.only(top: 8.0),
                  child: Text('UoM: ', style: TextStyle(fontSize: 24)),
                ),
                Expanded(
                  child: Wrap(
                    spacing: 8.0, // Jarak horizontal antar button
                    runSpacing: 8.0, // Jarak vertikal jika baris baru
                    children: uomOptions.map((uom) {
                      final isSelected = selectedUom == uom;
                      return ElevatedButton(
                        onPressed: () {
                          setState(() {
                            selectedUom = uom;
                          });
                        },
                        style: ElevatedButton.styleFrom(
                          // Warna background berubah jika terpilih
                          backgroundColor: isSelected
                              ? AppTheme
                                    .clicked // Warna saat aktif
                              : AppTheme.header, // Warna saat tidak aktif
                          foregroundColor: isSelected
                              ? Colors.white
                              : Colors.black,
                          shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(8),
                          ),
                        ),
                        child: Text(uom, style: TextStyle(fontSize: 18)),
                      );
                    }).toList(),
                  ),
                ),
              ],
            ),
            SizedBox(height: 24),
            TextField(
              controller: _remarksController,
              keyboardType: TextInputType.text,
              decoration: InputDecoration(
                label: Text('Remark', style: TextStyle(fontSize: 24)),
              ),
            ),
            SizedBox(height: 32),
            Center(
              child: ElevatedButton(
                onPressed: () async {
                  // Ambil nilai dari controller
                  String finalPartNum = _partNumController.text;
                  String finalName = _nameController.text;
                  String finalBrand = _brandController.text;
                  String finalRemarks = _remarksController.text;

                  // Validasi sederhana
                  if (finalPartNum.isEmpty ||
                      finalName.isEmpty ||
                      finalBrand.isEmpty) {
                    showDialog(
                      context: context,
                      builder: (context) => PopUp(
                        status: 'error',
                        title: 'Informasi produk kurang, mohon lengkapi:',
                        description: [
                          if (finalPartNum.isEmpty) 'Part Number',
                          if (finalName.isEmpty) 'Nama Produk',
                          if (finalBrand.isEmpty) 'Merk produk',
                        ].join(', '),
                      ),
                    );
                    return;
                  }
                  try {
                    final List<Map<String, dynamic>> safe = await _checkProduct(
                      finalPartNum,
                    );

                    if (widget.product == null && safe.isNotEmpty) {
                      productId =
                          int.tryParse(safe[0]['product_id'].toString()) ?? 0;
                      quantity +=
                          int.tryParse(safe[0]['stock_quantity'].toString()) ??
                          0;
                    }

                    if (widget.product == null && safe.isEmpty) {
                      Map<String, dynamic> data = {
                        "part_number": finalPartNum,
                        "product_name": finalName,
                        "brand": finalBrand,
                        "stock_quantity": quantity,
                        "unit": selectedUom,
                        "remarks": finalRemarks,
                      };
                      await _service.postData("product_services.php", data);
                    } else {
                      Map<String, dynamic> data = {
                        "product_id": productId,
                        "part_number": finalPartNum,
                        "product_name": finalName,
                        "brand": finalBrand,
                        "stock_quantity": quantity,
                        "unit": selectedUom,
                        "remarks": finalRemarks,
                      };
                      await _service.updateData("product_services.php", data);
                    }

                    if (context.mounted) {
                      await showDialog(
                        context: context,
                        builder: (context) => PopUp(
                          status: 'success',
                          title: 'Data Berhasil Diperbarui!',
                        ),
                      );
                      Navigator.pop(context);
                    }
                  } catch (e) {
                    if (context.mounted) {
                      await showDialog(
                        context: context,
                        builder: (context) => PopUp(
                          status: 'error',
                          title: 'Error : Gagal menyimpan data.',
                          description: 'Error: $e',
                        ),
                      );
                    }
                  }
                },
                child: const Text('Submit'),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Future<List<Map<String, dynamic>>> _checkProduct(String partNum) async {
    List<Map<String, dynamic>> data;

    data = await _service.getAllSearchedData("product_services.php", partNum);
    return data;
  }
}
