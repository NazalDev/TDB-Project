import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

// import appbar function
import 'package:tric_duta_bangsa/widgets/purchase_order/purchase_order_details.dart';
import 'package:tric_duta_bangsa/widgets/random/status_pop_up.dart';

class PurchaseOrderForm extends StatefulWidget {
  const PurchaseOrderForm({super.key});

  @override
  State<PurchaseOrderForm> createState() => _PurchaseOrderFormState();
}

class _PurchaseOrderFormState extends State<PurchaseOrderForm> {
  final ApiServices _services = ApiServices();

  // variables
  bool showCompanyInfo = false;
  bool showSiteInfo = false;

  // controller
  final TextEditingController _dateController = TextEditingController();
  final TextEditingController _companyName = TextEditingController();
  final TextEditingController _siteName = TextEditingController();
  final TextEditingController _siteAdress = TextEditingController();
  final TextEditingController _poNo = TextEditingController();
  String selectedId = '';

  List<DropdownMenuEntry<String>> siteEntries = [];
  List<Map<String, dynamic>> _allData = [];

  @override
  void initState() {
    super.initState();
    _loadSites(); // Fetch data when screen starts
  }

  //add site data into drop down
  Future<void> _loadSites() async {
    try {
      // Assuming your ProductService has a method like getAllProducts()
      // that returns a List of Maps or Models
      final List<Map<String, dynamic>> data = await _services.getAllData(
        "site_services.php",
      );

      setState(() {
        _allData = data;

        siteEntries = data.map((site) {
          return DropdownMenuEntry<String>(
            value: site['site_id'].toString(), // Use ID as value
            label:
                "${site['company_short_name']} ${site['site_name']}", // Use Name as label
          );
        }).toList();
      });
    } catch (e) {
      PopUp(status: 'error', title: 'Gagal Memuat Produk');
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
            // Site Information Section
            Row(
              children: [
                Text('Site       : ', style: TextStyle(fontSize: 24)),
                DropdownMenu(
                  label: const Text('Nama Site'),
                  enableFilter: true,
                  enableSearch: true,
                  onSelected: (value) {
                    setState(() {
                      final selectedSite = _allData.firstWhere(
                        (site) => site['site_id'].toString() == value,
                        orElse: () => {},
                      );

                      _companyName.text = selectedSite['company_name']
                          .toString();
                      _siteName.text = selectedSite['site_name'].toString();
                      _siteAdress.text = selectedSite['site_address']
                          .toString();
                      selectedId = selectedSite['site_id'].toString();
                    });
                  },
                  dropdownMenuEntries: siteEntries,
                ),
              ],
            ),
            SizedBox(height: 32),

            // Site Information Section
            Center(
              child: Text(
                'Informasi Site',
                style: TextStyle(fontSize: 32, fontWeight: FontWeight.normal),
              ),
            ),
            TextField(
              controller: _companyName,
              enabled: false,
              decoration: InputDecoration(
                label: Text('Nama Perusahaan', style: TextStyle(fontSize: 24)),
              ),
            ),
            SizedBox(height: 32),
            TextField(
              controller: _siteName,
              enabled: false,
              decoration: InputDecoration(
                label: Text('Nama Site', style: TextStyle(fontSize: 24)),
              ),
            ),
            SizedBox(height: 32),
            TextField(
              controller: _siteAdress,
              enabled: false,
              decoration: InputDecoration(
                label: Text('Alamat Site', style: TextStyle(fontSize: 24)),
              ),
            ),

            SizedBox(height: 32),

            //PO Information Section
            Center(
              child: Text(
                'Informasi Purchase Order',
                style: TextStyle(fontSize: 32),
              ),
            ),
            SizedBox(height: 32),
            TextField(
              controller: _poNo,
              keyboardType: TextInputType.text,
              decoration: InputDecoration(
                label: Text(
                  'Nomor Purchase Order',
                  style: TextStyle(fontSize: 24),
                ),
              ),
            ),
            SizedBox(height: 24),
            TextField(
              controller: _dateController,
              decoration: InputDecoration(
                labelText: 'TANGGAL PO',
                prefixIcon: Icon(Icons.calendar_today),
              ),
              readOnly: true,
              onTap: () {
                _selectDate();
              },
            ),
            SizedBox(height: 32),
            Center(
              child: ElevatedButton(
                onPressed: () async {
                  Map<String, dynamic> data = {
                    "po_number": _poNo.text,
                    "po_date": _dateController.text,
                    "site_id": selectedId,
                  };
                  final response = await _services.postData(
                    "purchase_order_services.php",
                    data,
                  );

                  if (response['status'] == 'success') {
                    final poNumber = response['po_number'];
                    if (context.mounted) {
                      Navigator.pushReplacement(
                        context,
                        MaterialPageRoute(
                          builder: (context) =>
                              PurchaseOrderDetail(poNumber: poNumber ?? ''),
                        ),
                      );
                    }
                  }
                  if (context.mounted) {
                    showDialog(
                      context: context,
                      builder: (context) => PopUp(
                        status: response['status'],
                        title: response['status'] == 'success'
                            ? 'Purchase Order Created Successfully'
                            : 'Failed to Create Purchase Order',
                        description: response['status'] == 'success'
                            ? null
                            : response['message'] ?? 'Unknown error',
                      ),
                    );
                  }
                },
                style: ButtonStyle(),
                child: const Text('Submit'),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Future<void> _selectDate() async {
    DateTime initialDate;

    if (_dateController.text.isNotEmpty) {
      initialDate = DateTime.parse(_dateController.text);
    } else {
      initialDate = DateTime.now();
    }

    DateTime? picked = await showDatePicker(
      context: context,
      initialDate: initialDate,
      firstDate: DateTime(2000),
      lastDate: DateTime(2100),
    );

    if (picked != null) {
      setState(() {
        _dateController.text = picked.toString().split(" ")[0];
      });
    }
  }
}
