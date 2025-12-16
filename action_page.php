<!doctype html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>案件查詢系統</title>

    <style>
        .container {
            max-width: 1200px;
            margin: 30px auto;
            padding: 0 20px;
        }

        h2.section-title {
            font-size: 22px;
            margin-bottom: 20px;
            color: #1976d2;
        }

        /* 卡片式連結區 */
        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 20px;
        }

        .card {
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            padding: 16px;
            display: flex;
            flex-direction: column;
            gap: 12px;
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }

        .card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
        }

        .card h3 {
            font-size: 16px;
            color: #444;
            border-left: 4px solid #1976d2;
            padding-left: 8px;
            margin-bottom: 8px;
        }

        .link-btn {
            display: block;
            padding: 10px;
            border-radius: 6px;
            background: #f0f4ff;
            color: #1976d2;
            text-decoration: none;
            font-weight: 500;
            transition: background 0.15s ease;
        }

        .link-btn:hover {
            background: #dce8ff;
        }
    </style>
</head>

<body>
    <?php include 'header.php'; ?>

    <div class="container">
        <hr style="margin: 40px 0;">

        <!-- 功能列表 -->
        <h2 class="section-title">功能列表</h2>
        <div class="grid">
            <div class="card">
                <h3>案件管理</h3>
                <a href="./case.php" class="link-btn">宏謙案件查詢</a>
                <a href="./case_old.php" class="link-btn">舊案件查詢</a>
                <a href="#" class="link-btn">查詢日誌</a>
                <a href="./acl_users.php" class="link-btn">權限管理</a>
            </div>
            <div class="card">
                <h3>統計報表</h3>
                <a href="./report_general.php" class="link-btn">一般報表</a>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script></script>
</body>

</html>