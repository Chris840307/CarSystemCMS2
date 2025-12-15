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
        .grid-top {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-bottom: 20px;
        }

        .grid-bottom {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 20px;
        }

        .card {
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

        /* 兩欄 */
        .card-links-2col {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 6px 12px;
        }

        @media (max-width: 768px) {
            .grid-top {
                grid-template-columns: 1fr;
            }

            .grid-bottom {
                grid-template-columns: 1fr;
            }

            .card-links-2col {
                grid-template-columns: 1fr;
            }
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
                ['label' => '一般報表', 'url' => 'report_general.php', 'color' => 'text-dark'],
                ['label' => '週報', 'url' => '', 'color' => 'text-dark']
            ];
            include 'breadcrumb.php';
            ?>

            <div class="grid-top">
                <div class="card">
                    <h3>缺漏號-手開單</h3>
                    <a href="./report/week/manual.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/week/manual.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/week/manual.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/week/manual.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/week/manual.php?organize=all" class="link-btn">全局</a>
                </div>
                <div class="card">
                    <h3>缺漏號-電子罰單</h3>
                    <a href="./report/week/ele_fine.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/week/ele_fine.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/week/ele_fine.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/week/ele_fine.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/week/ele_fine.php?organize=all" class="link-btn">全局</a>
                </div>
                <div class="card">
                    <h3>交通隊</h3>
                    <a href="./report/week/officers_list.php" class="link-btn">交通隊員警舉發清冊</a>
                    <a href="./report/week/performance.php" class="link-btn">績效新表</a>
                    <a href="./report/week/M49190.php" class="link-btn">王道遠(M49190)</a>
                </div>
            </div>
            <div class="grid-bottom">
                <div class="card">
                    <h3>整理交通秩序(惡性違規)</h3>
                    <div class="card-links-2col">
                        <a href="./report/statistics_record.php?type=week&period=current" class="link-btn">週報-當期</a>
                        <a href="./report/statistics_record.php?type=week&period=total" class="link-btn">週報-累計</a>
                        <a href="./report/statistics_public.php?type=week&period=current" class="link-btn">民眾檢舉-當期</a>
                        <a href="./report/statistics_public.php?type=week&period=total" class="link-btn">民眾檢舉-累計</a>
                        <a href="./report/statistics_technology_law.php?type=week&period=current" class="link-btn">科技執法-當期</a>
                        <a href="./report/statistics_technology_law.php?type=week&period=total" class="link-btn">科技執法-累計</a>
                        <a href="./report/statistics_remove.php?type=week&period=current" class="link-btn">剔除檢舉-當期</a>
                        <a href="./report/statistics_remove.php?type=week&period=total" class="link-btn">剔除檢舉-累計</a>
                        <a href="./report/statistics_police.php?type=week&period=current" class="link-btn">員警舉發-當期</a>
                        <a href="./report/statistics_police.php?type=week&period=total" class="link-btn">員警舉發-累計</a>
                    </div>
                </div>
                <div class="card">
                    <h3>其他</h3>
                    <div class="card-links-2col">
                        <a href="./report/double_parking.php?type=week" class="link-btn">並排停車舉發統計</a>
                        <a href="./report/large_car.php?type=week" class="link-btn">取締大型車</a>
                        <a href="./report/technology_law.php?type=week" class="link-btn">取締科技執法</a>
                        <a href="./report/drunk_driving.php?type=week" class="link-btn">取締酒駕</a>
                        <a href="./report/exhaust_pipe.php?type=week" class="link-btn">取締排氣管</a>
                        <a href="./report/plate_clean_up.php?type=week" class="link-btn">取締淨牌</a>
                        <a href="./report/no_yielding_to_pedestrians.php?type=week" class="link-btn">取締路口不停讓行人</a>
                        <!--與月報欄位不同-->
                        <a href="./report/week/dangerous_driving.php" class="link-btn">防制危險駕車績效成果表</a>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script></script>
</body>

</html>