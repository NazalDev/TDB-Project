import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

class Information {
  // Use 'static' so you don't have to call 'InformationHelper helper = InformationHelper();'
  // You can just call 'InformationHelper.buildHeader(...)'
  static Widget buildHeaderInfo({
    required List<dynamic> data,
    required Widget child,
  }) {
    if (data.isEmpty) {
      return const Center(child: CircularProgressIndicator());
    }

    return Container(
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(
        // Using Theme.of(context) makes it work in both Dark and Light mode automatically
        color: AppTheme.buttonNormal.withValues(alpha: 0.5),
        borderRadius: BorderRadius.circular(8),
        border: Border.all(color: AppTheme.textPrimary),
      ),
      child: child,
    );
  }

  static Widget infoRow(String label, String value, {TextStyle? labelStyle}) {
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
          Flexible(
            child: Text(
              value,
              textAlign: TextAlign.right,
              overflow: TextOverflow.ellipsis,
              style: TextStyle(color: AppTheme.textPrimary),
            ),
          ),
        ],
      ),
    );
  }
}
