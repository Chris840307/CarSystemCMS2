<?php
header("Content-Type: application/json");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

if ($method === "POST") {

    $data = json_decode(file_get_contents("php://input"), true);

    // 建立 where 條件
    $where = [];
    $params = [];

    if (!empty($data["form_No"])) {
        $where[] = "form_No LIKE ?";
        $params[] = "%" . $data["form_No"] . "%";
    }

    if (!empty($data["plate_number"])) {
        $where[] = "plate_number LIKE ?";
        $params[] = "%" . $data["plate_number"] . "%";
    }

    if (!empty($data["violation_date"])) {
        $where[] = "violation_date = ?";
        $params[] = $data["violation_date"];
    }

    if (!empty($data["department"])) {
        $where[] = "department = ?";
        $params[] = $data["department"];
    }

    if (!empty($data["officer"])) {
        $where[] = "officer LIKE ?";
        $params[] = "%" . $data["officer"] . "%";
    }

    $sql = "SELECT *
            FROM case_old";
    // 若有條件才增加 WHERE
    if (count($where) > 0) {
        $sql .= " WHERE " . implode(" AND ", $where);
    }
    $sql .= " ORDER BY id DESC";

    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

    echo json_encode([
        "returnCode" => 200,
        "data" => $rows
    ]);
} else {
    http_response_code(405);
    echo json_encode(["error" => "Method not allowed"]);
}
