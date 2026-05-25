<?php
$host = "localhost";
$user = "root";
$pass = ""; // Password MySQL kamu
$db   = "tric_duta_bangsa";

$conn = new mysqli($host, $user, $pass, $db);

if ($conn->connect_error) {
    die(json_encode(["status" => "error", "message" => "Koneksi Gagal"]));
}

?>