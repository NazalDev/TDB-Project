import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

// import appbar function

import 'package:tric_duta_bangsa/widgets/random/status_pop_up.dart';

class SiteForm extends StatefulWidget {
  const SiteForm({super.key});

  @override
  State<SiteForm> createState() => _SiteFormState();
}

class _SiteFormState extends State<SiteForm> {
  final ApiServices _services = ApiServices();

  bool _openCompanyInfo = false;

  //controller
  final TextEditingController _companyName = TextEditingController();
  final TextEditingController _companyShortName = TextEditingController();
  final TextEditingController _companyNPWP = TextEditingController();
  final TextEditingController _companyAddress = TextEditingController();
  final TextEditingController _siteName = TextEditingController();
  final TextEditingController _siteAddress = TextEditingController();
  String _companyId = '';

  List<DropdownMenuEntry<String>> companyEntries = [];
  List<Map<String, dynamic>> _allData = [];

  @override
  void initState() {
    super.initState();
    _loadCompanies(); // Fetch data when screen starts
  }

  //add site data into drop down
  Future<void> _loadCompanies() async {
    try {
      // Assuming your ProductService has a method like getAllProducts()
      // that returns a List of Maps or Models
      final List<Map<String, dynamic>> data = await _services.getAllData(
        "company_services.php",
      );

      setState(() {
        _allData = data;

        companyEntries = data.map((company) {
          return DropdownMenuEntry<String>(
            value: company['company_short_name'].toString(), // Use ID as value
            label: company['company_short_name']
                .toString(), // Use Name as label
          );
        }).toList();

        // Add the "Add New Product" option at the end or beginning
        companyEntries.add(
          const DropdownMenuEntry(
            value: 'openForm',
            label: 'Tambah Perusahaan Baru',
          ),
        );
      });
    } catch (e) {
      if (!mounted) return;

      showDialog(
        context: context,
        builder: (context) => PopUp(
          status: 'error',
          title: 'Gagal load data',
          description: e.toString(),
        ),
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
          shrinkWrap: true,
          children: [
            // Company Information Section
            Row(
              children: [
                Text('Perusahaan : ', style: TextStyle(fontSize: 24)),
                DropdownMenu(
                  label: const Text('Nama Perusahaan'),
                  enableFilter: true,
                  enableSearch: true,
                  onSelected: (value) {
                    setState(() {
                      if (value == 'openForm') {
                        _companyName.clear();
                        _companyShortName.clear();
                        _companyNPWP.clear();
                        _companyAddress.clear();
                        _openCompanyInfo = true;
                      } else if (value != null && value != 'openForm') {
                        final selectedCompany = _allData.firstWhere(
                          (company) =>
                              company['company_short_name'].toString() == value,
                          orElse: () => {},
                        );

                        _companyName.text = selectedCompany['company_name']
                            .toString();
                        _companyShortName.text =
                            selectedCompany['company_short_name'].toString();
                        _companyNPWP.text = selectedCompany['company_npwp']
                            .toString();
                        _companyAddress.text =
                            selectedCompany['company_address'].toString();
                        _companyId = selectedCompany['company_id'].toString();
                        _openCompanyInfo = false;
                      }
                    });
                  },
                  dropdownMenuEntries: companyEntries,
                ),
              ],
            ),

            SizedBox(height: 16),
            Column(
              children: [
                Text(
                  'Informasi Perusahaan',
                  style: TextStyle(fontSize: 32, fontWeight: FontWeight.normal),
                ),
                TextField(
                  controller: _companyName,
                  keyboardType: TextInputType.text,
                  enabled: _openCompanyInfo,
                  decoration: InputDecoration(
                    label: Text(
                      'Nama Perusahaan',
                      style: TextStyle(fontSize: 24),
                    ),
                  ),
                ),
                SizedBox(height: 32),
                TextField(
                  controller: _companyShortName,
                  keyboardType: TextInputType.text,
                  enabled: _openCompanyInfo,
                  decoration: InputDecoration(
                    label: Text(
                      'Nama Singkat Perusahaan',
                      style: TextStyle(fontSize: 24),
                    ),
                  ),
                ),
                SizedBox(height: 32),
                TextField(
                  controller: _companyNPWP,
                  keyboardType: TextInputType.text,
                  enabled: _openCompanyInfo,
                  decoration: InputDecoration(
                    label: Text(
                      'NPWP Perusahaan',
                      style: TextStyle(fontSize: 24),
                    ),
                  ),
                ),
                SizedBox(height: 32),
                TextField(
                  controller: _companyAddress,
                  keyboardType: TextInputType.text,
                  enabled: _openCompanyInfo,
                  decoration: InputDecoration(
                    label: Text(
                      'Alamat Perusahaan',
                      style: TextStyle(fontSize: 24),
                    ),
                  ),
                ),
              ],
            ),

            SizedBox(height: 32),
            Column(
              children: [
                Text(
                  'Informasi Site',
                  style: TextStyle(fontSize: 32, fontWeight: FontWeight.normal),
                ),
                TextField(
                  controller: _siteName,
                  keyboardType: TextInputType.text,
                  decoration: InputDecoration(
                    label: Text('Nama Site', style: TextStyle(fontSize: 24)),
                  ),
                ),
                SizedBox(height: 32),
                TextField(
                  controller: _siteAddress,
                  keyboardType: TextInputType.text,
                  decoration: InputDecoration(
                    label: Text('Alamat Site', style: TextStyle(fontSize: 24)),
                  ),
                ),
                SizedBox(height: 32),
                Center(
                  child: ElevatedButton(
                    onPressed: () async {
                      Map<String, dynamic> dataC = {
                        "company_name": _companyName.text,
                        "company_short_name": _companyShortName.text,
                        "company_npwp": _companyNPWP.text,
                        "company_address": _companyAddress.text,
                      };

                      final Map<String, dynamic> res;

                      if (_openCompanyInfo) {
                        res = await _services.postData(
                          "company_services.php",
                          dataC,
                        );
                      } else {
                        res = {"status": "success", "company_id": _companyId};
                      }

                      if (res['status'] == "success") {
                        Map<String, dynamic> dataS = {
                          "company_id": res['company_id'].toString(),
                          "site_name": _siteName.text,
                          "site_address": _siteAddress.text,
                        };

                        var res2 = await _services.postData(
                          "site_services.php",
                          dataS,
                        );

                        if (!context.mounted) return;
                        Navigator.pop(context);

                        if (res2['status'] == "success") {
                          showDialog(
                            context: context,
                            builder: (context) => PopUp(
                              status: 'success',
                              title: 'Data Site Berhasil Disimpan',
                            ),
                          );
                        } else {
                          showDialog(
                            context: context,
                            builder: (context) => PopUp(
                              status: 'error',
                              title: 'Gagal',
                              description: 'Gagal menyimpan data site',
                            ),
                          );
                        }
                      }
                    },
                    style: ButtonStyle(),
                    child: const Text('Submit'),
                  ),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }
}
