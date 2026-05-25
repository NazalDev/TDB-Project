<?php
// product_services.php

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: POST, GET, DELETE, PUT, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, Access-Control-Allow-Headers, Authorization, X-Requested-With");
header('Content-Type: application/json');
file_put_contents('log.txt', $_SERVER['REQUEST_METHOD'] . " request received at " . date('Y-m-d H:i:s') . "\n", FILE_APPEND);
include 'db_config.php';

$method = $_SERVER['REQUEST_METHOD'];

if ($method == 'POST') {
    // Menerima input dari Flutter
    $input = json_decode(file_get_contents('php://input'), true);
    
    $stmt = $conn->prepare("INSERT INTO company_sites (company_id, site_name, site_address) VALUES (?, ?, ?)");
    $stmt->bind_param("sss", 
        $input['company_id'], 
        $input['site_name'], 
        $input['site_address']
    );

    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Produk berhasil ditambah"]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
    
}

if ($method == 'GET') {
    $search = isset($_GET['search']) ? $_GET['search'] : '';
    
    if ($search !== '') {
        // Mode search
        $stmt = $conn->prepare("SELECT cs.site_id, cs.site_name, cs.site_address, c.company_name, c.company_short_name FROM company_sites cs JOIN companies c ON cs.company_id = c.company_id WHERE c.company_name LIKE ? or cs.site_name LIKE ?");
        $likeTerm = "%" . $search . "%";
        $stmt->bind_param("ss", $likeTerm, $likeTerm);
        $stmt->execute();
        $result = $stmt->get_result();
    } else {
        // mode get all
        $result = $conn->query("SELECT cs.site_id, cs.site_name, cs.site_address, c.company_name, c.company_short_name FROM company_sites cs JOIN companies c ON cs.company_id = c.company_id");
    }

    $data = [];
    while($row = $result->fetch_assoc()) { 
        $data[] = $row; 
    }
    echo json_encode($data);
}

if ($method == 'DELETE') {
    $id = isset($_GET['site_id']) ? intval($_GET['site_id']) : 0;
    
    if ($id > 0) {
        $stmt = $conn->prepare("DELETE FROM company_sites WHERE site_id = ?");
        $stmt->bind_param("i", $id);
        $stmt->execute();
    } else {
        http_response_code(400);
        echo json_encode(["status" => "error", "message" => "Invalid product ID"]);
    }
}

$conn->close();
?>