<?php
header("Content-Type: application/json; charset=utf-8");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method === "POST") {

    $data = json_decode(file_get_contents("php://input"), true);

    if (!isset($data['user_id'], $data['permission'], $data['value'])) {
        http_response_code(400);
        echo json_encode(["success" => false, "message" => "參數錯誤"]);
        exit;
    }

    $userId = intval($data['user_id']);
    $permissionCode = $data['permission'];
    $value = $data['value'] ? 1 : 0;

    try {
        $pdo->beginTransaction();

        // 先查 permission_id
        $stmt = $pdo->prepare("SELECT id FROM permissions WHERE code = ?");
        $stmt->execute([$permissionCode]);
        $perm = $stmt->fetch(PDO::FETCH_ASSOC);

        if (!$perm) {
            throw new Exception("找不到該權限");
        }

        $permissionId = intval($perm['id']);

        if ($value === 1) {
            // 插入 (如果不存在)
            $stmt = $pdo->prepare("
                INSERT IGNORE INTO user_permissions (user_id, permission_id) 
                VALUES (?, ?)
            ");
            $stmt->execute([$userId, $permissionId]);
        } else {
            // 刪除
            $stmt = $pdo->prepare("
                DELETE FROM user_permissions 
                WHERE user_id = ? AND permission_id = ?
            ");
            $stmt->execute([$userId, $permissionId]);
        }

        $pdo->commit();
        echo json_encode(["success" => true]);

    } catch (Exception $e) {
        $pdo->rollBack();
        http_response_code(500);
        echo json_encode(["success" => false, "message" => $e->getMessage()]);
    }

} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
