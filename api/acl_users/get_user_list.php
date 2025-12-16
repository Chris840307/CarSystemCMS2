<?php
header("Content-Type: application/json; charset=utf-8");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method === "GET") {
    // 抓所有users與他們的權限
    $sql = "SELECT 
                A.id, A.acc, A.name, A.department,
                C.id AS permission_id, C.name AS permission_name
            FROM users A
            LEFT JOIN user_permissions B ON A.id = B.user_id
            LEFT JOIN permissions C ON B.permission_id = C.id
            ORDER BY C.id
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
                "acc" => $row['acc'],
                "name" => $row['name'],
                "department" => $row['department'],
                "permissions" => []
            ];
        }

        if ($row['permission_id'] !== null) {
            $users[$uid]['permissions'][] = [
                "id" => $row['permission_id'],
                "name" => $row['permission_name']
            ];
        }
    }

    echo json_encode([
        "returnCode" => 200,
        "data" => array_values($users)
    ]);
} else if ($method === "POST") {
    $data = json_decode(file_get_contents("php://input"), true);

    $sql = "SELECT 
                A.id, A.acc, A.name, A.department,
                C.id AS permission_id, C.name AS permission_name
            FROM users A
            LEFT JOIN user_permissions B ON A.id = B.user_id
            LEFT JOIN permissions C ON B.permission_id = C.id
            WHERE A.id = ?
            ORDER BY C.id";
    $stmt = $pdo->prepare($sql);
    $stmt->execute([$data['user_id'] ?? '']);

    $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // 整理成前端格式
    $users = [];
    foreach ($rows as $row) {
        $uid = $row['id'];

        if (!isset($users[$uid])) {
            $users[$uid] = [
                "id" => $uid,
                "acc" => $row['acc'],
                "name" => $row['name'],
                "department" => $row['department'],
                "permissions" => []
            ];
        }

        if ($row['permission_id'] !== null) {
            $users[$uid]['permissions'][] = [
                "id" => $row['permission_id'],
                "name" => $row['permission_name']
            ];
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
