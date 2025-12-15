<?php
$host = "117.56.148.80";
$user = "root";
$pass = "2u6u/ru8";
$dbname = "CMS2";

try {
    $pdo = new PDO(
        "mysql:host=$host;dbname=$dbname;charset=utf8",
        $user,
        $pass,
        [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]
    );
} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode(["error" => "Database connection failed", "detail" => $e->getMessage()]);
    exit;
}
