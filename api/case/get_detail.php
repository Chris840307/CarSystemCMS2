<?php
header("Content-Type: application/json");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method == "POST") {
    $data = json_decode(file_get_contents("php://input"), true);

    $sql = "SELECT *
            FROM case_old
            ORDER BY id DESC";
    $stmt = $pdo->prepare($sql);
    $stmt->execute([$data["id"]]);
    $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

    echo json_encode([
        "returnCode" => 200,
        "data" => $rows
    ]);
} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
