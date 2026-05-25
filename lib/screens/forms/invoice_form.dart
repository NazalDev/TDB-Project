import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

class InvoiceDialog extends StatefulWidget {
  final Function(DateTime dueDate, double tax, String currency) onConfirm;

  const InvoiceDialog({super.key, required this.onConfirm});

  @override
  State<InvoiceDialog> createState() => _InvoiceDialogState();
}

class _InvoiceDialogState extends State<InvoiceDialog> {
  final DateTime _selectedDate = DateTime.now();
  final TextEditingController _dateController = TextEditingController(
    text: "30",
  );
  final TextEditingController _taxController = TextEditingController(
    text: "11",
  );
  final TextEditingController _currencyController = TextEditingController(
    text: "IDR",
  );

  @override
  void dispose() {
    _taxController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      title: const Text(
        "Create Invoice",
        style: TextStyle(fontWeight: FontWeight.bold, fontSize: 32),
        textAlign: TextAlign.center,
      ),
      backgroundColor: AppTheme.box,
      content: Column(
        mainAxisSize: MainAxisSize.min,
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            "Term of Payment:",
            style: TextStyle(fontWeight: FontWeight.bold),
          ),
          const SizedBox(height: 8),

          // Date Picker Button
          TextField(
            controller: _dateController,
            keyboardType: TextInputType.number,
            decoration: const InputDecoration(
              border: OutlineInputBorder(),
              suffixText: "Hari",
            ),
          ),

          const SizedBox(height: 20),
          const Text("PPN:", style: TextStyle(fontWeight: FontWeight.bold)),
          const SizedBox(height: 8),

          // Tax Input
          TextField(
            controller: _taxController,
            keyboardType: TextInputType.number,
            decoration: const InputDecoration(
              border: OutlineInputBorder(),
              suffixText: "%",
              hintText: "11",
            ),
          ),

          const SizedBox(height: 20),
          const Text(
            "Currency:",
            style: TextStyle(fontWeight: FontWeight.bold),
          ),
          const SizedBox(height: 8),

          TextField(
            controller: _currencyController,
            textCapitalization:
                TextCapitalization.characters, // Forces uppercase
            decoration: const InputDecoration(
              border: OutlineInputBorder(),
              hintText: "IDR",
            ),
          ),

          const SizedBox(height: 20),
        ],
      ),
      actions: [
        Row(
          mainAxisAlignment: MainAxisAlignment.center,
          crossAxisAlignment: CrossAxisAlignment.center,
          children: [
            TextButton(
              onPressed: () => Navigator.pop(context),
              style: ElevatedButton.styleFrom(overlayColor: Colors.red),
              child: const Text("Cancel"),
            ),
            ElevatedButton(
              style: ElevatedButton.styleFrom(
                backgroundColor: AppTheme.greenMedium,
              ),
              onPressed: () {
                double taxValue = double.tryParse(_taxController.text) ?? 11.0;
                widget.onConfirm(
                  _selectedDate.add(
                    Duration(days: int.tryParse(_dateController.text) ?? 30),
                  ),
                  taxValue,
                  _currencyController.text,
                );
                Navigator.pop(context);
              },
              child: const Text(
                "Create Invoice",
                style: TextStyle(color: Colors.black),
              ),
            ),
          ],
        ),
      ],
    );
  }
}
