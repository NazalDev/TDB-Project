// ignore_for_file: avoid_print

import 'package:http/http.dart' as http;
import 'dart:convert';

class ApiServices {
  //final String baseUrl = "http://192.168.7.7/tdb_api";
  final String baseUrl = "http://localhost/tdb_api";

  Future<List<Map<String, dynamic>>> getAllData(String endPoint) async {
    final url = Uri.parse('$baseUrl/$endPoint');

    try {
      final response = await http.get(url);

      if (response.statusCode == 200) {
        // Mengubah string JSON dari PHP menjadi List di Flutter
        List<dynamic> data = json.decode(response.body);
        return data.map((item) => item as Map<String, dynamic>).toList();
      } else {
        throw Exception("Gagal memuat data");
      }
    } catch (e) {
      print("Error API: $e");
      return [];
    }
  }

  Future<List<Map<String, dynamic>>> getAllSearchedData(
    String endPoint,
    String query,
  ) async {
    // Menambahkan query parameter ke URL
    final url = Uri.parse('$baseUrl/$endPoint?search=$query');

    try {
      final response = await http.get(url);

      if (response.statusCode == 200) {
        List<dynamic> data = json.decode(response.body);
        return data.map((item) => item as Map<String, dynamic>).toList();
      } else {
        throw Exception("Gagal mencari data");
      }
    } catch (e) {
      print("Error API Search: $e");
      return [];
    }
  }

  Future<Map<String, dynamic>> postData(
    String endPoint,
    Map<String, dynamic> body,
  ) async {
    try {
      final response = await http.post(
        Uri.parse('$baseUrl/$endPoint'),
        headers: {"Content-Type": "application/json"},
        body: json.encode(body),
      );

      return json.decode(response.body);
    } catch (e) {
      return {"status": 'error', "message": e.toString()};
    }
  }

  Future<Map<String, dynamic>> deleteData(String endPoint, String query) async {
    final url = Uri.parse('$baseUrl/$endPoint?search=$query');

    try {
      final response = await http.delete(url);
      return json.decode(response.body);
    } catch (e) {
      return {"status": 'error', "message": "Error: $e", "statusCode": 0};
    }
  }

  Future<Map<String, dynamic>> updateData(
    String endPoint,
    Map<String, dynamic> body,
  ) async {
    final url = Uri.parse('$baseUrl/$endPoint');

    try {
      final response = await http.put(
        url,
        headers: {"Content-Type": "application/json"},
        body: json.encode(body),
      );

      if (response.statusCode == 200 || response.statusCode == 201) {
        return json.decode(response.body);
      } else {
        return json.decode(response.body);
      }
    } catch (e) {
      return {"status": 'error', "message": e.toString()};
    }
  }
}
