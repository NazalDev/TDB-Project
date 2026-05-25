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
    $searchTerm = isset($_GET['search']) ? $_GET['search'] : '';
    
    $stmt = $conn->prepare("SELECT ii.quantity, p.part_number, p.unit, p.is_active, ii.price, ii.total  FROM invoice_items ii JOIN products p USING (product_id) WHERE ii.invoice_id LIKE ?");
    $stmt->bind_param("s", $searchTerm);
    $stmt->execute();
    $result = $stmt->get_result();
    
    $data = [];
    while($row = $result->fetch_assoc()) { 
        $data[] = $row; 
    }
    echo json_encode($data);
}

$conn->close();
?>