import 'package:flutter/material.dart';
import 'package:tric_duta_bangsa/utils/app_style.dart';

class PopUp extends StatelessWidget {
  final String title;
  final String status;
  final String? description;

  const PopUp({
    super.key,
    required this.status,
    required this.title,
    this.description,
  });

  @override
  Widget build(BuildContext context) {
    final isSuccess = status == 'success';
    final isError = status == 'error';

    final bgColor = isSuccess
        ? AppTheme.accentGreen
        : isError
        ? AppTheme.accentRed
        : AppTheme.clicked;

    final icon = isSuccess
        ? Icons.check_circle_rounded
        : isError
        ? Icons.cancel_rounded
        : Icons.info_rounded;

    return AlertDialog(
      backgroundColor: bgColor,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
      content: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          Icon(icon, size: 72, color: Colors.white),
          const SizedBox(height: 12),
          Text(
            title,
            textAlign: TextAlign.center,
            style: const TextStyle(
              fontFamily: 'Fredoka One', // add to pubspec.yaml
              fontSize: 22,
              fontWeight: FontWeight.bold,
              color: Colors.white,
            ),
          ),
          if (isError && description != null) ...[
            const SizedBox(height: 8),
            Text(
              description!,
              textAlign: TextAlign.center,
              style: const TextStyle(
                fontFamily: 'Fredoka One',
                fontSize: 15,
                color: Colors.white70,
              ),
            ),
          ],
          const SizedBox(height: 20),
          ElevatedButton(
            style: ElevatedButton.styleFrom(
              backgroundColor: Colors.white,
              foregroundColor: bgColor,
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(12),
              ),
            ),
            onPressed: () => Navigator.pop(context),
            child: const Text(
              'OK',
              style: TextStyle(
                fontFamily: 'Fredoka One',
                fontWeight: FontWeight.bold,
              ),
            ),
          ),
        ],
      ),
    );
  }
}
