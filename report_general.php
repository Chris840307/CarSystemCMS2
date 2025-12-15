<!doctype html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>案件查詢系統</title>

    <style>
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
            max-width: 300px;
            background: #fff;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, .08);
            padding: 16px;
            display: flex;
            flex-direction: column;
            gap: 10px;
            transition: .2s;
        }

        .card:hover {
            transform: translateY(-3px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, .14);
        }

        .card h3 {
            font-size: 18px;
            border-left: 4px solid #1976d2;
            padding-left: 8px;
            color: #444;
        }

        .link-btn {
            display: block;
            background: #f0f4ff;
            border-radius: 6px;
            padding: 10px;
            margin-bottom: 6px;
            color: #1976d2;
            text-decoration: none;
            font-weight: 500;
        }

        .link-btn:hover {
            background: #dce8ff;
        }
    </style>
</head>

<body>
    <?php include 'header.php'; ?>

    <div class="container-fluid">
        <div class="main-container">
            <?php
            $breadcrumbs = [
                ['label' => '首頁', 'url' => 'action_page.php', 'color' => 'text-secondary'],
                ['label' => '一般報表', 'url' => '', 'color' => 'text-dark']
            ];
            include 'breadcrumb.php';
            ?>
            <hr style="margin: 20px 0;">

            <h2 class="section-title">月報</h2>
            <div class="grid">
                <div class="card">
                    <h3>115</h3>
                    <a href="./report_general_month.php?month=1" class="link-btn">1月</a>
                    <!-- <a href="./report_general_month.php?month=2" class="link-btn">2月</a> -->
                </div>
            </div>

            <hr style="margin: 40px 0;">

            <h2 class="section-title">週報</h2>
            <div class="grid">
                <div class="card">
                    <h3>115</h3>
                    <a href="./report_general_week.php?week=1" class="link-btn">第一週</a>
                    <!-- <a href="./report_general_week.php?week=1" class="link-btn">第二週</a> -->
                </div>
            </div>

            <hr style="margin: 40px 0;">

            <h2 class="section-title">其他</h2>
            <div class="grid">
                <div class="card">
                    <h3>其他</h3>
                    <a href="./report/other/sign_sheet.php" class="link-btn">標示單清冊</a>
                    <a href="./report/other/no_sign_sheet.php" class="link-btn">未達清冊</a>
                </div>
            </div>

            <hr style="margin: 40px 0;">

            <h2 class="section-title">獎勵金</h2>
            <div class="grid">
                <div class="card">
                    <h3>其他</h3>
                    <a href="./report/bonus/.php" class="link-btn">交通隊交通事故件數</a>
                    <a href="./report/bonus/.php" class="link-btn">交通隊各項統計</a>
                </div>
                <div class="card">
                    <h3>外勤員警獎勵金印領清冊</h3>
                    <a href="./report/bonus/receipt_list.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/bonus/receipt_list.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/bonus/receipt_list.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/bonus/receipt_list.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/bonus/receipt_list.php?organize=safety" class="link-btn">保安隊</a>
                </div>
                <div class="card">
                    <h3>惡性違規核對表</h3>
                    <a href="./report/bonus/vicious_violation.php?name=呂元皓&id=E60680" class="link-btn">呂元皓(E60680)</a>
                </div>
                <div class="card">
                    <h3>獎勵金點數核對表</h3>
                    <a href="./report/bonus/bonus_points.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/bonus/bonus_points.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/bonus/bonus_points.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/bonus/bonus_points.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/bonus/bonus_points.php?organize=garrison" class="link-btn">警備隊</a>
                </div>
                <div class="card">
                    <h3>交整時數表</h3>
                    <a href="./report/bonus/hours.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/bonus/hours.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/bonus/hours.php?organize=three" class="link-btn">三分局</a>
                </div>
                <div class="card">
                    <h3>排氣管改裝取締報表</h3>
                    <a href="./report/bonus/exhaust_pipe.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/bonus/exhaust_pipe.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/bonus/exhaust_pipe.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/bonus/exhaust_pipe.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/bonus/exhaust_pipe.php?organize=all" class="link-btn">全局</a>
                </div>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script></script>
</body>

</html>