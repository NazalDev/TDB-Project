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
        $stmt = $conn->prepare("SELECT po.po_number, po.po_date, po.po_status, cs.site_name, cs.site_address, c.company_name, c.company_short_name FROM purchase_orders po JOIN company_sites cs USING (site_id) JOIN companies c USING (company_id) WHERE po.po_number LIKE ? AND po.is_deleted = 0");
        $likeTerm = "%" . $searchTerm . "%";
        $stmt->bind_param("s", $likeTerm);
        $stmt->execute();
        $result = $stmt->get_result();
    } else {
        // Mode Biasa: Mengambil semua data produk
        $result = $conn->query("SELECT po.po_number, po.po_date, po.po_status, cs.site_name, cs.site_address, c.company_name, c.company_short_name FROM purchase_orders po JOIN company_sites cs USING (site_id) JOIN companies c USING (company_id) WHERE po.is_deleted = 0");
    }

    $data = [];
    while($row = $result->fetch_assoc()) { 
        $data[] = $row; 
    }
    echo json_encode($data);

} elseif ($method == 'POST') {
    // Menerima input dari Flutter
    $input = json_decode(file_get_contents('php://input'), true);
    
    $stmt = $conn->prepare("INSERT INTO purchase_orders (po_number, po_date, site_id) VALUES (?, ?, ?)");
    $stmt->bind_param("ssi", 
        $input['po_number'], 
        $input['po_date'], 
        $input['site_id']
    );

    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Purchase Order berhasil ditambah", "po_number" => $input['po_number']]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
    
} elseif ($method == 'DELETE') {
    $id = isset($_GET['search']) ? $_GET['search'] : '';

    if ($id !== '') {
        $stmt = $conn->prepare("UPDATE purchase_orders SET is_deleted = 1, deleted_at = NOW() WHERE po_number = ?");
        $stmt->bind_param("s", $id);
        
        if ($stmt->execute()) {
            if ($stmt->affected_rows > 0) {
                http_response_code(200);
                echo json_encode(["status" => "success", "message" => "Deleted successfully"]);

            } else {
                http_response_code(404);
                echo json_encode(["status" => "error", "message" => "No Po found with Number: $id"]);
            }
        } else {
            http_response_code(400);
            echo json_encode(["status" => "error", "message" => $stmt->error]);
        }
    } else {
        http_response_code(400);
        echo json_encode(["status" => "error", "message" => "Invalid Po Number"]);
    }
    exit;
    
} elseif ($method == 'PUT') {

    $input = json_decode(file_get_contents("php://input"), true);

    if (!isset($input['po_number'])) {
        echo json_encode(["status" => "error", "message" => "ID wajib diisi"]);
        exit;
    }

    $stmt = $conn->prepare(
        "UPDATE purchase_orders 
         SET po_date=?, site_id=? 
         WHERE po_number=?"
    );

    $stmt->bind_param(
        "sis",
        $input['po_date'],
        $input['site_id'],
        $input['po_number'],
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