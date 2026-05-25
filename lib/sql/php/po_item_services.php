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
    
    $stmt = $conn->prepare("SELECT poi.po_item_id, poi.quantity, poi.unit_price, poi.total, p.product_id, p.part_number, p.product_name, p.unit, p.stock_quantity, p.is_active FROM purchase_order_items poi JOIN products p USING (product_id) WHERE poi.po_number LIKE ? AND poi.is_deleted = 0");
    $stmt->bind_param("s", $searchTerm);
    $stmt->execute();
    $result = $stmt->get_result();
    
    $data = [];
    while($row = $result->fetch_assoc()) { 
        $data[] = $row; 
    }
    echo json_encode($data);
}

if ($method == 'POST') {
    $input = json_decode(file_get_contents('php://input'), true);

    if (!isset($input['po_number'], $input['product_id'], $input['quantity'], $input['unit_price'])) {
        echo json_encode(["status" => "error", "message" => "Missing required fields"]);
        exit;
    }

    $total = $input['quantity'] * $input['unit_price'];

    // Check if the product already exists in this purchase order
    $checkStmt = $conn->prepare("
        SELECT po_item_id, is_deleted 
        FROM purchase_order_items 
        WHERE po_number = ? AND product_id = ?
    ");
    $checkStmt->bind_param("ss", $input['po_number'], $input['product_id']);
    $checkStmt->execute();
    $existing = $checkStmt->get_result()->fetch_assoc();

    if ($existing) {
        if ($existing['is_deleted'] == 0) {
            // Row exists and is active → block the insert
            echo json_encode(["status" => "duplicate", "message" => "Produk sudah ada di purchase order ini"]);
            exit;
        } else {
            // Row exists but was soft-deleted → restore it with new values
            $stmt = $conn->prepare("
                UPDATE purchase_order_items 
                SET quantity = ?, unit_price = ?, total = ?, is_deleted = 0, deleted_at = NULL
                WHERE po_item_id = ?
            ");
            $stmt->bind_param("ssss", $input['quantity'], $input['unit_price'], $total, $existing['po_item_id']);
        }
    } else {
        // Row doesn't exist → insert it
        $stmt = $conn->prepare("
            INSERT INTO purchase_order_items (po_number, product_id, quantity, unit_price, total) 
            VALUES (?, ?, ?, ?, ?)
        ");
        $stmt->bind_param("sssss", 
            $input['po_number'], 
            $input['product_id'], 
            $input['quantity'], 
            $input['unit_price'],
            $total
        );
    }

    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Produk berhasil ditambah"]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
}

if ($method == 'DELETE') {
    $id = isset($_GET['search']) ? $_GET['search'] : '';

    if ($id !== '') {
        $stmt = $conn->prepare("UPDATE purchase_order_items SET is_deleted = 1, deleted_at = NOW() WHERE po_item_id = ?");
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
    
}

$conn->close();
?>