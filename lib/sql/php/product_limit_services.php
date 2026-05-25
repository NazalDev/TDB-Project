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
        $stmt = $conn->prepare("SELECT DISTINCT p.* FROM products p
                INNER JOIN purchase_order_items poi USING (product_id)
                INNER JOIN delivery_orders USING (po_number) WHERE do_id LIKE ?");
        $stmt->bind_param("s", $searchTerm);
        $stmt->execute();
        $result = $stmt->get_result();
    }

    $data = [];
    while($row = $result->fetch_assoc()) { 
        $data[] = $row; 
    }
    echo json_encode($data);

}
$conn->close();
?>