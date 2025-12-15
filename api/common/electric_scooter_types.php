<?php
header("Content-Type: application/json");
require_once "../db.php";

$method = $_SERVER["REQUEST_METHOD"];

switch ($method) {
    case "GET":
        $sql = "SELECT * FROM electric_scooter_types";
        $stmt = $pdo->prepare($sql);
        $stmt->execute();
        $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

        $result = [];
        foreach ($rows as $row) {
            $id = $row['id'];
            if (!isset($result[$id])) {
                $result[$id] = [
                    'key' => $row['value'],
                    'value' => $row['text']
                ];
            }
        }

         echo json_encode([
            "returnCode" => 200,
            "data" => array_values($result)
        ]);
        break;
    default:
        http_response_code(405);
        echo json_encode(["error" => "Method not allowed"]);
}
