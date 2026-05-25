import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:tric_duta_bangsa/sql/api_services.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';
import 'dart:async';

import 'package:tric_duta_bangsa/screens/forms/invoice_form.dart';
import 'package:tric_duta_bangsa/widgets/invoice/invoice_details.dart';
import 'package:tric_duta_bangsa/widgets/random/button.dart';
import 'package:tric_duta_bangsa/widgets/random/pop_up_form.dart';
import 'package:tric_duta_bangsa/widgets/random/status_pop_up.dart';

class DeliveryOrderDetails extends StatefulWidget {
  final String doNumber;
  final String doId;
  final bool openForInvoice;
  const DeliveryOrderDetails({
    super.key,
    required this.doNumber,
    required this.doId,
    this.openForInvoice = false,
  });

  @override
  State<DeliveryOrderDetails> createState() => _DeliveryOrderDetailsState();
}

class _DeliveryOrderDetailsState extends State<DeliveryOrderDetails> {
  final ApiServices _services = ApiServices();
  final Button _button = Button();

  List<dynamic> productToAdd = [];
  List<dynamic> productToUpdate = [];
  List<dynamic> productToRemove = [];
  List<dynamic> deliveryOrder = [];
  List<dynamic> poItems = [];
  List<dynamic> doItems = [];

  bool isLoading = false;
  bool hiddenInvoiceDialog = false;
  bool _hasChanged = false;

  final TextEditingController _notesController = TextEditingController();
  final TextEditingController _qtyController = TextEditingController(text: '0');

  @override
  void dispose() {
    _notesController.dispose(); // Always clean up controllers
    super.dispose();
  }

  @override
  void initState() {
    super.initState();
    _initData();

    if (widget.openForInvoice) {
      WidgetsBinding.instance.addPostFrameCallback((_) {
        _showInvoiceDialog();
      });
    }
  }

  Future<void> _initData() async {
    setState(() => isLoading = true);

    final doResult = await _services.getAllSearchedData(
      "delivery_order_services.php",
      widget.doNumber,
    );

    if (doResult.isEmpty) {
      setState(() => isLoading = false);
      return;
    }

    final doData = doResult[0];
    final bool isPoDeleted =
        doData['po_is_deleted'] == 1 ||
        doData['po_is_deleted'] == '1' ||
        doData['po_is_deleted'] == true;

    // Fetch PO items and DO items in parallel using the returned doData
    final results = await Future.wait([
      _services.getAllSearchedData("po_item_services.php", doData['po_number']),
      _services.getAllSearchedData("do_item_services.php", widget.doId),
    ]);

    final fetchedPoItems = results[0];
    final fetchedDoItems = results[1];

    for (final item in fetchedDoItems) {
      item['quantity'] = item['qty_shipped']?.toString() ?? '0';
    }

    final poMap = {
      for (final item in fetchedPoItems) item['product_id'].toString(): item,
    };

    for (final item in fetchedDoItems) {
      final match = poMap[item['product_id'].toString()];
      if (match != null) {
        item['quantity'] = match['quantity'];
      }
    }

    // Filter out products already in DO from the PO list
    final existingIds = fetchedDoItems
        .map((item) => item['product_id'])
        .toSet();
    final filteredPoItems = fetchedPoItems
        .where((item) => !existingIds.contains(item['product_id']))
        .toList();

    setState(() {
      deliveryOrder = doResult;
      poItems = filteredPoItems;
      doItems = fetchedDoItems;
      hiddenInvoiceDialog = isPoDeleted;
      isLoading = false;
    });
  }

  @override
  Widget build(BuildContext context) {
    return PopScope(
      canPop: !_hasChanged,
      onPopInvokedWithResult: (didPop, result) {
        if (didPop) return;
        if (_hasChanged) {
          savePopUp(context, true);
        }
      },
      child: Scaffold(
        backgroundColor: AppTheme.box,
        appBar: AppBar(
          title: Text(
            // ignore: unnecessary_string_interpolations
            "${widget.doNumber}",
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
                  const SizedBox(width: 8),
                  if (!hiddenInvoiceDialog) _buildInvoiceButton(),
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
                  child: Row(
                    children: [
                      productList(Icons.add, true, poItems),
                      VerticalDivider(color: AppTheme.accentRed, thickness: 2),
                      productList(Icons.remove, false, doItems),
                    ],
                  ),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Future<dynamic> savePopUp(BuildContext context, bool close) {
    return FormPopUp.show(
      (context),
      title: 'Simpan Perubahan?',
      maxWidth: 200,
      showCloseButton: false,
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        crossAxisAlignment: CrossAxisAlignment.center,
        children: [
          _button.buildButton(
            onPressed: () {
              Navigator.pop(context);
              if (!close) return;
              Navigator.pop(context);
            },
            text: 'Tidak',
          ),
          _button.buildButton(
            onPressed: () {
              _saveDoItems();
              _deleteDoItems();
              Navigator.pop(context);
              if (!close) return;
              Navigator.pop(context);
            },
            text: 'Iya',
          ),
        ],
      ),
    );
  }

  Future<void> _saveDoItems() async {
    if (productToAdd.isEmpty && productToUpdate.isEmpty) return;

    setState(() => isLoading = true);
    List<String> errors = [];

    // POST — new items
    for (final item in productToAdd) {
      final result = await _services.postData('do_item_services.php', {
        "do_id": widget.doId,
        "po_item_id": item['po_item_id'],
        "quantity": item['qty_shipped'],
      });
      if (result['status'] != 'success') {
        errors.add(item['part_number'] ?? 'unknown');
      }
    }

    // PUT — existing items with updated qty
    for (final item in productToUpdate) {
      final result = await _services.updateData('do_item_services.php', {
        "do_item_id": item['do_item_id'],
        "quantity": item['qty_shipped'],
      });
      if (result['status'] != 'success') {
        errors.add(item['part_number'] ?? 'unknown');
      }
    }

    setState(() {
      isLoading = false;
      _hasChanged = false; // reset change flag after saving
    });

    if (!mounted) return;
    // Show result feedback
    if (errors.isEmpty) {
      showDialog(
        context: context,
        builder: (context) =>
            PopUp(status: 'success', title: 'Semua Data Berhasil Disimpan'),
      );
    } else {
      showDialog(
        context: context,
        builder: (context) => PopUp(
          status: 'error',
          title: 'Gagal Menyimpan Data:',
          description: errors.join(', '),
        ),
      );
    }
  }

  Future<void> _deleteDoItems() async {
    if (productToRemove.isEmpty) return;

    setState(() => isLoading = true);

    // Loop through each item and post one by one
    List<String> errors = [];
    for (final item in productToRemove) {
      final result = await _services.deleteData(
        'do_item_services.php',
        item['do_item_id'].toString(),
      );

      if (result['status'] != 'success') {
        errors.add(item['part_number'] ?? 'unknown');
      }
    }

    setState(() {
      isLoading = false;
      productToRemove.clear();
    });
  }

  Expanded productList(IconData? icon, bool isAdd, List<dynamic> data) {
    return Expanded(
      child: ListView.separated(
        itemCount: data.length,
        separatorBuilder: (context, index) => SizedBox(height: 8),
        itemBuilder: (context, index) {
          final items = data[index];
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
                          items['part_number'] ?? 'N/A',
                          style: TextStyle(
                            fontWeight: FontWeight.w500,
                            fontSize: 16,
                          ),
                        ),
                        Text(items['product_name'] ?? 'N/A'),
                      ],
                    ),

                    Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        // Edit qty button — only shown on the DO (right) list
                        if (!isAdd)
                          IconButton(
                            onPressed: () => _showEditQtyPopup(items),
                            style: ElevatedButton.styleFrom(
                              backgroundColor: AppTheme.clicked,
                              foregroundColor: Colors.white,
                            ),
                            icon: const Icon(Icons.edit),
                          ),
                        // Add: opens qty popup | Remove: instant
                        _buildButton(
                          () async {
                            if (isAdd) {
                              await _showAddQtyPopup(items);
                            } else {
                              _addRemoveFunc(false, items);
                            }
                          },
                          icon,
                          isAdd,
                        ),
                      ],
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
                        'Stock : ${items['stock_quantity']} ${items['unit']}',
                      ),
                      if (isAdd)
                        Text(
                          'Qty to Ship : ${items['quantity']} ${items['unit']}',
                        ),
                      if (!isAdd)
                        if (!isAdd)
                          Text(
                            'Qty to Ship : ${items['qty_shipped']} ${items['unit']}',
                          ),
                    ],
                  ),
                ),
              ],
            ),
          );
        },
      ),
    );
  }

  /// Opens the qty popup for ADDING an item from the PO list to the DO list.
  Future<void> _showAddQtyPopup(dynamic items) async {
    int qty = 0;
    _qtyController.text = '0';

    await FormPopUp.show(
      context,
      maxWidth: 200,
      showCloseButton: false,
      title: 'Set Qty Do',

      child: StatefulBuilder(
        builder: (context, setPopupState) {
          return Row(
            mainAxisAlignment: MainAxisAlignment.spaceEvenly,
            crossAxisAlignment: CrossAxisAlignment.center,
            children: [
              IconButton(
                onPressed: () {
                  if (qty > 0) {
                    setPopupState(() {
                      qty--;
                      _qtyController.text = qty.toString();
                    });
                  }
                },
                icon: const Icon(Icons.remove),
              ),
              SizedBox(
                width: 100,
                child: TextField(
                  controller: _qtyController,
                  textAlign: TextAlign.center,
                  keyboardType: TextInputType.number,
                  inputFormatters: [FilteringTextInputFormatter.digitsOnly],
                  onChanged: (value) {
                    if (value.isEmpty) {
                      setPopupState(() => qty = 0);
                    } else {
                      final parsed = int.parse(value);
                      final maxStock = int.parse(
                        items['stock_quantity'].toString(),
                      );
                      final maxQty = int.parse(items['quantity'].toString());
                      if (parsed > maxStock || parsed > maxQty) {
                        final maxxed = maxStock > maxQty ? maxQty : maxStock;
                        _qtyController.text = maxxed.toString();
                        setPopupState(() => qty = maxxed);
                      } else {
                        setPopupState(() => qty = parsed);
                      }
                    }
                  },
                ),
              ),
              IconButton(
                onPressed: () {
                  final maxStock = int.parse(
                    items['stock_quantity'].toString(),
                  );
                  final maxQty = int.parse(items['quantity'].toString());
                  if (qty < maxStock && qty < maxQty) {
                    setPopupState(() {
                      qty++;
                      _qtyController.text = qty.toString();
                    });
                  }
                },
                icon: const Icon(Icons.add),
              ),
            ],
          );
        },
      ),
    );

    if (qty > 0) {
      items['qty_shipped'] = qty.toString();
      _addRemoveFunc(true, items);
    }
  }

  /// Opens the qty popup for EDITING an existing item already in the DO list.
  Future<void> _showEditQtyPopup(dynamic items) async {
    int qty = int.tryParse(items['qty_shipped'].toString()) ?? 0;
    _qtyController.text = qty.toString();

    await FormPopUp.show(
      context,
      maxWidth: 200,
      showCloseButton: false,
      title: 'Edit Qty Do',

      child: StatefulBuilder(
        builder: (context, setPopupState) {
          return Row(
            mainAxisAlignment: MainAxisAlignment.spaceEvenly,
            crossAxisAlignment: CrossAxisAlignment.center,
            children: [
              IconButton(
                onPressed: () {
                  if (qty > 1) {
                    setPopupState(() {
                      qty--;
                      _qtyController.text = qty.toString();
                    });
                  }
                },
                icon: const Icon(Icons.remove),
              ),
              SizedBox(
                width: 100,
                child: TextField(
                  controller: _qtyController,
                  textAlign: TextAlign.center,
                  keyboardType: TextInputType.number,
                  inputFormatters: [FilteringTextInputFormatter.digitsOnly],
                  onChanged: (value) {
                    if (value.isEmpty) {
                      setPopupState(() => qty = 1);
                    } else {
                      final parsed = int.parse(value);
                      final maxStock = int.parse(
                        items['stock_quantity'].toString(),
                      );
                      final maxQty = int.parse(items['quantity'].toString());
                      if (parsed > maxStock || parsed > maxQty) {
                        final maxxed = maxStock > maxQty ? maxQty : maxStock;
                        _qtyController.text = maxxed.toString();
                        setPopupState(() => qty = maxxed);
                      } else if (parsed < 1) {
                        _qtyController.text = '1';
                        setPopupState(() => qty = 1);
                      } else {
                        setPopupState(() => qty = parsed);
                      }
                    }
                  },
                ),
              ),
              IconButton(
                onPressed: () {
                  final max = int.parse(items['stock_quantity'].toString());
                  final maxQty = int.parse(items['quantity'].toString());
                  if (qty < max && qty < maxQty) {
                    setPopupState(() {
                      qty++;
                      _qtyController.text = qty.toString();
                    });
                  }
                },
                icon: const Icon(Icons.add),
              ),
            ],
          );
        },
      ),
    );

    // Commit the edit
    setState(() {
      items['qty_shipped'] = qty.toString();
      _hasChanged = true;
    });

    if (items['do_item_id'] != null) {
      productToUpdate.removeWhere(
        (e) => e['do_item_id'] == items['do_item_id'],
      );
      productToUpdate.add(items);
    }
  }

  void _addRemoveFunc(bool isAdd, items) {
    return setState(() {
      if (isAdd) {
        if (items['do_item_id'] != null) {
          productToRemove.remove(items);
        } else {
          items['qty_shipped'] = _qtyController.text;
          productToAdd.add(items);
        }

        doItems.add(items);
        poItems.remove(items);
        _hasChanged = true;
      } else {
        if (items['do_item_id'] != null) {
          productToRemove.add(items);
        } else {
          productToAdd.remove(items);
        }

        doItems.remove(items);
        poItems.add(items);
        _hasChanged = true;
      }
    });
  }

  Widget _buildButton(VoidCallback onPressed, IconData? icon, bool isAdd) {
    return IconButton(
      onPressed: onPressed,
      style: ElevatedButton.styleFrom(
        backgroundColor: isAdd ? AppTheme.accentGreen : AppTheme.accentRed,
        foregroundColor: Colors.white,
      ),
      icon: Icon(icon),
    );
  }

  Widget _buildInvoiceButton() {
    return ElevatedButton.icon(
      onPressed: () async {
        if (_hasChanged) {
          await FormPopUp.show(
            (context),
            title: 'Simpan Perubahan?',
            maxWidth: 200,
            showCloseButton: false,
            child: Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              crossAxisAlignment: CrossAxisAlignment.center,
              children: [
                _button.buildButton(
                  onPressed: () => Navigator.pop(context),
                  text: 'Tidak',
                ),
                _button.buildButton(
                  onPressed: () {
                    _saveDoItems();
                    Navigator.pop(context);
                  },
                  text: 'Iya',
                ),
              ],
            ),
          );
        }
        _showInvoiceDialog();
      },
      style: ElevatedButton.styleFrom(
        backgroundColor: AppTheme.clicked,
        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
      ),
      icon: const Icon(Icons.description, color: Colors.white),
      label: const Text(
        'Generate Invoice',
        style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold),
      ),
    );
  }

  void _showInvoiceDialog() {
    showDialog(
      context: context,
      builder: (context) => InvoiceDialog(
        onConfirm: (dueDate, ppn, curr) async {
          // This calls your existing API logic
          final Map<String, dynamic> data = {
            "do_id": widget.doId,
            "due_date": dueDate.toIso8601String(),
            "ppn": ppn.toString(),
            "currency": curr.toString(),
          };

          final res = await _services.postData("invoice_services.php", data);

          if (context.mounted) {
            if (res['status'] == 'success') {
              Navigator.push(
                context,
                MaterialPageRoute(
                  builder: (context) => InvoiceDetails(
                    invoiceId: res['invoice_id'].toString(),
                    invoiceNumber: res['invoice_number'].toString(),
                  ),
                ),
              );
            }

            showDialog(
              context: context,
              builder: (context) => PopUp(
                status: res['status'],
                title: 'Invoice berhasil dibuat',
              ),
            );
          }
        },
      ),
    );
  }

  Widget _buildHeaderInfo() {
    if (deliveryOrder.isEmpty) {
      return const Center(child: CircularProgressIndicator());
    }

    final data = deliveryOrder[0];

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
          _infoRow("Date:", data['do_date'] ?? 'N/A'),
          _infoRow("Status:", '-'),
          _infoRow("Notes:", data['notes'] ?? 'N/A'),
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
}
