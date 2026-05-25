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
    
    $stmt = $conn->prepare("INSERT INTO companies (company_name, company_short_name, company_npwp, company_address) VALUES (?, ?, ?, ?)");
    $stmt->bind_param("ssss", 
        $input['company_name'], 
        $input['company_short_name'], 
        $input['company_npwp'], 
        $input['company_address']
    );

    if ($stmt->execute()) {
        $lastId = $conn->insert_id;
        echo json_encode(["status" => "success", "message" => "Produk berhasil ditambah", "company_id" => $lastId]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
    
}

if ($method == 'GET') {
    $result = $conn->query("SELECT * FROM companies");

    $data = [];
    while($row = $result->fetch_assoc()) { 
        $data[] = $row; 
    }
    echo json_encode($data);
}

$conn->close();
?>