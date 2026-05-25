import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

class Button {
  Widget buildButton({
    required VoidCallback onPressed,
    String? text,
    IconData? icon,
    BuildContext? context,
  }) {
    return Container(
      decoration: BoxDecoration(
        border: Border.all(color: AppTheme.textPrimary, width: 3),
        borderRadius: BorderRadius.circular(8),
      ),
      child: ElevatedButton.icon(
        onPressed: onPressed,
        style: ElevatedButton.styleFrom(
          backgroundColor: AppTheme.buttonNormal,
          shape: const RoundedRectangleBorder(borderRadius: BorderRadius.zero),
        ),
        icon: Icon(icon, color: Colors.black),
        label: Text(
          text ?? '',
          style: TextStyle(color: Colors.black, fontWeight: FontWeight.bold),
        ),
      ),
    );
  }

  Row functionalButton({
    VoidCallback? onPressedV,
    VoidCallback? onPressedE,
    VoidCallback? onPressedD,
    VoidCallback? onPressedM,
    String? mlabel,
  }) {
    return Row(
      mainAxisSize: MainAxisSize.min,
      mainAxisAlignment: MainAxisAlignment.center,
      crossAxisAlignment: CrossAxisAlignment.center,
      children: [
        if (onPressedV != null) ...[
          IconButton.filled(
            icon: Icon(Icons.remove_red_eye_outlined),
            onPressed: onPressedV,
            style: IconButton.styleFrom(
              backgroundColor: Colors.greenAccent,
              foregroundColor: Colors.black,
              highlightColor: Colors.green.withValues(alpha: 0.4),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(8),
              ),
            ),
          ),
        ],

        if (onPressedE != null) ...[
          SizedBox(width: 4),
          IconButton.filled(
            icon: Icon(Icons.edit),
            onPressed: onPressedE,
            style: IconButton.styleFrom(
              backgroundColor: Colors.yellowAccent,
              foregroundColor: Colors.black,
              highlightColor: Colors.yellow.withValues(alpha: 0.6),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(8),
              ),
            ),
          ),
        ],

        if (onPressedD != null) ...[
          SizedBox(width: 4),
          IconButton.filled(
            icon: Icon(Icons.delete),
            onPressed: onPressedD,
            style: IconButton.styleFrom(
              backgroundColor: Colors.redAccent,
              foregroundColor: Colors.black,
              highlightColor: Colors.red.withValues(alpha: 0.4),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(8),
              ),
            ),
          ),
        ],

        if (onPressedM != null) ...[
          SizedBox(width: 4),
          ElevatedButton.icon(
            onPressed: onPressedM,
            label: Text(mlabel ?? '', style: TextStyle(fontSize: 16)),
            icon: Icon(Icons.upload_file_outlined),
            style: IconButton.styleFrom(
              iconSize: 26,
              fixedSize: Size.fromHeight(40),
              backgroundColor: Colors.blueAccent,
              foregroundColor: Colors.black,
              highlightColor: Colors.blue.withValues(alpha: 0.4),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(8),
              ),
            ),
          ),
        ],
      ],
    );
  }
}
