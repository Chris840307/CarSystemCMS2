<?php
header("Content-Type: application/json; charset=utf-8");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method === "GET") {
    // 抓所有users與他們的權限
    $sql = "SELECT 
                A.id, A.acc, A.name, A.department,
                C.status AS permission_code
            FROM users A
            LEFT JOIN user_permissions B ON A.id = B.user_id
            LEFT JOIN permissions C ON B.permission_id = C.id
            ORDER BY A.id
        ";

    $stmt = $pdo->prepare($sql);
    $stmt->execute();
    $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // 整理成前端格式
    $users = [];
    foreach ($rows as $row) {
        $uid = $row['id'];

        if (!isset($users[$uid])) {
            $users[$uid] = [
                "id" => $uid,
                "account" => $row['account'],
                "name" => $row['name'],
                "department" => $row['department'],
                "unit" => $row['unit'],
                "permissions" => []
            ];
        }

        if (!empty($row['permission_code'])) {
            $users[$uid]['permissions'][] = $row['permission_code'];
        }
    }

    echo json_encode([
        "returnCode" => 200,
        "data" => array_values($users)
    ]);
} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
