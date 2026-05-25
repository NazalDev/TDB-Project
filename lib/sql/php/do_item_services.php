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
    
    $stmt = $conn->prepare("SELECT doi.do_item_id, doi.qty_shipped, p.product_id, p.part_number, p.product_name, p.stock_quantity, p.unit, p.is_active  FROM delivery_order_items doi JOIN products p USING (product_id) WHERE doi.do_id LIKE ? AND doi.is_deleted = 0");
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
    // Menerima input dari Flutter
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (!isset($input['do_id'], $input['quantity'], $input['po_item_id'])) {
        echo json_encode(["status" => "error", "message" => "Missing required fields"]);
        exit;
    }
    
    // fetch price from po item
    $infoStmt = $conn->prepare("
        SELECT unit_price, product_id 
        FROM purchase_order_items        
        WHERE po_item_id = ?
    ");
    $infoStmt->bind_param("s", $input['po_item_id']);
    $infoStmt->execute();
    $info = $infoStmt->get_result()->fetch_assoc();
    
    if (!$info) {
        echo json_encode(["status" => "error", "message" => "PO item not found"]);
        exit;
    }
    $total = $input['quantity'] * $info['unit_price'];
    
    $checkStmt = $conn->prepare("
        SELECT do_item_id, is_deleted 
        FROM delivery_order_items 
        WHERE do_id = ? AND product_id = ?
    ");
    $checkStmt->bind_param("ss", $input['do_id'], $info['product_id']);
    $checkStmt->execute();
    $existing = $checkStmt->get_result()->fetch_assoc();

    if ($existing) {
        if ($existing['is_deleted'] == 0) {
            // Row exists and is active → block the insert
            echo json_encode(["status" => "duplicate", "message" => "Produk sudah ada di delivery order ini"]);
            exit;
        } else {
            // Row exists but was soft-deleted → restore it with new values
            $stmt = $conn->prepare("
                UPDATE delivery_order_items 
                SET qty_shipped = ?, unit_price = ?, total = ?, is_deleted = 0, deleted_at = NULL
                WHERE do_item_id = ?
            ");
            $stmt->bind_param("ssss", $input['quantity'], $info['unit_price'], $total, $existing['do_item_id']);
        }
    } else {
        // Row doesn't exist → insert it
        $stmt = $conn->prepare("
            INSERT INTO delivery_order_items (do_id, product_id, qty_shipped, unit_price, total) 
            VALUES (?, ?, ?, ?, ?)
        ");
        $stmt->bind_param("sssss", 
            $input['do_id'], 
            $info['product_id'], 
            $input['quantity'],    
            $info['unit_price'],
            $total 
        );
    }

    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Produk berhasil ditambah"]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
}

if ($method == 'PUT') {
    $input = json_decode(file_get_contents('php://input'), true);

    if (!isset($input['do_item_id'], $input['quantity'])) {
        echo json_encode(["status" => "error", "message" => "Missing required fields"]);
        exit;
    }

    // Recalculate total based on existing unit_price
    $priceStmt = $conn->prepare("SELECT unit_price FROM delivery_order_items WHERE do_item_id = ?");
    $priceStmt->bind_param("s", $input['do_item_id']);
    $priceStmt->execute();
    $priceRow = $priceStmt->get_result()->fetch_assoc();

    if (!$priceRow) {
        echo json_encode(["status" => "error", "message" => "Item not found"]);
        exit;
    }

    $total = $input['quantity'] * $priceRow['unit_price'];

    $stmt = $conn->prepare("
        UPDATE delivery_order_items 
        SET qty_shipped = ?, total = ? 
        WHERE do_item_id = ?
    ");
    $stmt->bind_param("sss", $input['quantity'], $total, $input['do_item_id']);

    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Quantity updated"]);
    } else {
        echo json_encode(["status" => "error", "message" => $stmt->error]);
    }
    exit;
}

if ($method == 'DELETE') {
    $id = isset($_GET['search']) ? $_GET['search'] : '';

    if ($id !== '') {
        $stmt = $conn->prepare("UPDATE delivery_order_items SET is_deleted = 1, deleted_at = NOW() WHERE do_item_id = ?");
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