import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

/// A fully generalized popup dialog that wraps any form widget.
///
/// Usage examples:
///
/// 1. Simple static form popup:
///    FormPopUp.show(context, title: 'Add Product', child: MyFormWidget());
///
/// 2. With a callback when dialog closes (e.g. to refresh a list):
///    FormPopUp.show(context, title: 'Add Site', child: SiteForm(), onClose: _reload);
///
/// 3. As a widget (e.g. inside showDialog manually):
///    showDialog(context: context, builder: (_) => FormPopUp(title: '...', child: ...));

class FormPopUp extends StatelessWidget {
  final String title;
  final Widget child;

  /// Optional: controls the max width of the dialog content.
  final double? maxWidth;

  /// Optional: controls the max height of the dialog content.
  final double? maxHeight;

  /// Optional: whether to show a close (X) button in the dialog title.
  final bool showCloseButton;

  const FormPopUp({
    super.key,
    required this.title,
    required this.child,
    this.maxWidth,
    this.maxHeight,
    this.showCloseButton = true,
  });

  /// Convenience static method to show the popup from anywhere.
  /// [onClose] is called with the result returned by Navigator.pop inside the child form.
  static Future<T?> show<T>(
    BuildContext context, {
    required String title,
    required Widget child,
    double? maxWidth,
    double? maxHeight,
    bool showCloseButton = true,
    VoidCallback? onClose,
  }) async {
    final result = await showDialog<T>(
      context: context,
      builder: (_) => FormPopUp(
        title: title,
        maxWidth: maxWidth,
        maxHeight: maxHeight,
        showCloseButton: showCloseButton,
        child: child,
      ),
    );
    if (onClose != null) onClose();
    return result;
  }

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      backgroundColor: AppTheme.box,
      title: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Expanded(
            child: Text(
              title.toUpperCase(),
              style: TextStyle(fontSize: 32, fontWeight: FontWeight.bold),
              textAlign: TextAlign.center,
            ),
          ),
          if (showCloseButton)
            IconButton(
              icon: const Icon(Icons.close),
              onPressed: () => Navigator.of(context).pop(),
              tooltip: 'Close',
            ),
        ],
      ),
      content: SizedBox(
        width: maxWidth ?? double.maxFinite,
        height: maxHeight, // null = auto height
        child: SingleChildScrollView(child: child),
      ),
    );
  }
}
