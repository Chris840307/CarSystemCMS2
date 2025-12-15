<?php
header("Content-Type: application/json");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method == "POST") {
    $data = json_decode(file_get_contents("php://input"), true);

    $sql = "SELECT * FROM `users` WHERE `acc` = ? LIMIT 1";
    $stmt = $pdo->prepare($sql);
    $stmt->execute([$data['account'] ?? '']);

    $row = $stmt->fetch(PDO::FETCH_ASSOC);

    if ($row && $row['pwd'] === $data['pwd']) {
        echo json_encode([
            "returnCode" => 200,
            "data" => [
                "user_id" => $row['id'],
                "name" => $row['name']
            ]
        ]);
    } else {
        http_response_code(400);
        echo json_encode(["message" => "帳號或密碼錯誤"]);
    }
} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
