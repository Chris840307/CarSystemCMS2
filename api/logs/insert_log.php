<?php
header("Content-Type: application/json; charset=utf-8");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method === "POST") {
    $input = json_decode(file_get_contents('php://input'), true);

    // 驗證必填欄位
    $requiredFields = ['username', 'searchr', 'detail'];
    foreach ($requiredFields as $field) {
        if (!isset($input[$field])) {
            http_response_code(400);
            echo json_encode(["error" => "$field 欄位缺失"]);
            exit;
        }
    }

    $sql = "INSERT INTO `search_log` (`username`, `ip`, `searchr`, `useragent`, `detail`) 
            VALUES (:username, :ip, :searchr, :useragent, :detail)";
    $stmt = $pdo->prepare($sql);
    try {
        $stmt->execute([
            ':username' => $input['username'],
            ':ip' => $_SERVER['REMOTE_ADDR'],
            ':searchr' => $input['searchr'],
            ':useragent' => $_SERVER['HTTP_USER_AGENT'],
            ':detail' => $input['detail'],
        ]);

        echo json_encode([
            "returnCode" => 200,
            "message" => "新增成功",
            "id" => $pdo->lastInsertId()
        ]);
    } catch (PDOException $e) {
        http_response_code(500);
        echo json_encode(["error" => "資料庫錯誤: " . $e->getMessage()]);
    }
} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
