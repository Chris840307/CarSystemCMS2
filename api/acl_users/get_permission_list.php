<?php
header("Content-Type: application/json; charset=utf-8");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method === "GET") {
    $sql = "SELECT id, `status`, `name` FROM permissions ORDER BY id";
    $stmt = $pdo->prepare($sql);
    $stmt->execute();
    $permissions = $stmt->fetchAll(PDO::FETCH_ASSOC);

    echo json_encode([
        "returnCode" => 200,
        "data" => $permissions
    ]);

} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
