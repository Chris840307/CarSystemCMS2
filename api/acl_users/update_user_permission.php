<?php
header("Content-Type: application/json; charset=utf-8");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];


if ($method === "POST") {
    $data = json_decode(file_get_contents("php://input"), true);

    if (!isset($data['user_id'], $data['permission_id'], $data['value'])) {
        http_response_code(400);
        echo json_encode(["success" => false, "message" => "參數錯誤"]);
        exit;
    }

    $userId = (int)$data['user_id'];
    $permissionId = (int)$data['permission_id'];
    $value = $data['value'];

    try {
        $pdo->beginTransaction();

        if ($value) {
            $stmt = $pdo->prepare("
            INSERT IGNORE INTO user_permissions (user_id, permission_id)
            VALUES (?, ?)
        ");
            $stmt->execute([$userId, $permissionId]);
        } else {
            $stmt = $pdo->prepare("
            DELETE FROM user_permissions
            WHERE user_id = ? AND permission_id = ?
        ");
            $stmt->execute([$userId, $permissionId]);
        }

        $pdo->commit();

        echo json_encode([
            "returnCode" => 200,
            "message" => 'success',
        ]);
    } catch (Exception $e) {
        $pdo->rollBack();
        http_response_code(500);
        echo json_encode(["message" => $e->getMessage()]);
    }
} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
