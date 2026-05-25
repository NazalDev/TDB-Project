<?php
// product_services.php

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: POST, GET, DELETE, PUT, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, Access-Control-Allow-Headers, Authorization, X-Requested-With");
header('Content-Type: application/json');
file_put_contents('log.txt', $_SERVER['REQUEST_METHOD'] . " request received at " . date('Y-m-d H:i:s') . "\n", FILE_APPEND);
include 'db_config.php';

$method = $_SERVER['REQUEST_METHOD'];

if ($method == 'GET') {
    // cek apakah ada parameter pencarian
    $searchTerm = isset($_GET['search']) ? $_GET['search'] : '';

    if ($searchTerm !== '') {
        // Mode Pencarian: Menggunakan LIKE untuk mencari part_number
        $stmt = $conn->prepare("SELECT * FROM products WHERE part_number LIKE ? AND is_active = 1");
        $likeTerm = "%" . $searchTerm . "%";
        $stmt->bind_param("s", $likeTerm);
        $stmt->execute();
        $result = $stmt->get_result();
    } else {
        // Mode Biasa: Mengambil semua data produk
        $result = $conn->query("SELECT * FROM products WHERE is_active = 1");
    }

    $data = [];
    while($row = $result->fetch_assoc()) { 
        $data[] = $row; 
    }
    echo json_encode($data);

} elseif ($method == 'POST') {
    // Menerima input dari Flutter
    $input = json_decode(file_get_contents('php://input'), true);
    
    $stmt = $conn->prepare("INSERT INTO products (part_number, product_name, brand, stock_quantity, unit, remarks) VALUES (?, ?, ?, ?, ?, ?)");
    $stmt->bind_param("sssiss", 
        $input['part_number'], 
        $input['product_name'], 
        $input['brand'], 
        $input['stock_quantity'], 
        $input['unit'], 
        $input['remarks']
    );

    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Produk berhasil ditambah"]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
    
} elseif ($method == 'DELETE') {
    $id = isset($_GET['search']) ? intval($_GET['search']) : 0;

    if ($id > 0) {
        $stmt = $conn->prepare("UPDATE products SET is_active = 0, deleted_at = NOW() WHERE product_id = ?");
        $stmt->bind_param("i", $id);
        $stmt->execute();
        
        if ($stmt->execute()) {
            if ($stmt->affected_rows > 0) {
                http_response_code(200);
                echo json_encode(["status" => "success", "message" => "Deleted successfully"]);
            } else {
                http_response_code(404);
                echo json_encode(["status" => "error", "message" => "No product found with ID: $id"]);
            }
        } else {
            http_response_code(400);
            echo json_encode(["status" => "error", "message" => $stmt->error]);
        }
    } else {
        http_response_code(400);
        echo json_encode(["status" => "error", "message" => "Invalid product ID"]);
    }
    exit;
    
} elseif ($method == 'PUT') {

    $input = json_decode(file_get_contents("php://input"), true);

    if (!isset($input['product_id'])) {
        echo json_encode(["status" => "error", "message" => "ID wajib diisi"]);
        exit;
    }

    $stmt = $conn->prepare(
        "UPDATE products 
         SET part_number=?, product_name=?, brand=?, stock_quantity=?, unit=?, remarks=? 
         WHERE product_id=?"
    );

    $stmt->bind_param(
        "sssissi",
        $input['part_number'],
        $input['product_name'],
        $input['brand'],
        $input['stock_quantity'],
        $input['unit'],
        $input['remarks'],
        $input['product_id']
    );
    $stmt->execute();
    
    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Produk berhasil diupdate"]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
}
$conn->close();
?>