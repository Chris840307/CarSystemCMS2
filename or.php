<?php
// Oracle 連線資訊
$host = '125.227.94.19';
$port = 1521;
$sid = 'orcl'; // 例如: ORCL
$username = 'traffic';
$password = 'joly902f';

// 建立連接字串 (TNS)
$tns = "(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = $host)(PORT = $port))
    (CONNECT_DATA =
        (SERVICE_NAME = $sid)
    )
  )";

// 建立連線
$conn = oci_pconnect($username, $password, $tns,'AL32UTF8');

if (!$conn) {
    $e = oci_error();
    echo "連接失敗: " . $e['message'];
} else {
    echo "連接成功！";
    // 這裡可以執行SQL操作
    //oci_close($conn);
}

$stid = oci_parse($conn, 'select * from Unitinfo');
oci_execute($stid);
oci_close($conn);
echo "<table border='0'>\n";
while ($row = oci_fetch_array($stid, OCI_ASSOC+OCI_RETURN_NULLS)) {
    echo "<tr>\n";
    foreach ($row as $item) {
        echo "    <td>" . ($item !== null ? htmlentities($item, ENT_QUOTES) : "&nbsp;") . "</td>\n";
    }
    echo "</tr>\n";
}
echo "</table>\n";
?>
