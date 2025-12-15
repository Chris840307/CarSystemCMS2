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
                ['label' => '月報', 'url' => '', 'color' => 'text-dark']
            ];
            include 'breadcrumb.php';
            ?>

            <div class="grid-top">
                <div class="card">
                    <h3>取締大型車</h3>
                    <a href="./report/ban_large_vehicles.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/ban_large_vehicles.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/ban_large_vehicles.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/ban_large_vehicles.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/ban_large_vehicles.php?organize=all" class="link-btn">全局</a>
                </div>
                <div class="card">
                    <h3>成果統計表(警察機關裁罰)</h3>
                    <a href="./report/achievement_police.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/achievement_police.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/achievement_police.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/achievement_police.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/achievement_police.php?organize=all" class="link-btn">全局</a>
                </div>
                <div class="card">
                    <h3>成果統計表(移公路主管機關裁罰)</h3>
                    <a href="./report/achievement_highway.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/achievement_highway.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/achievement_highway.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/achievement_highway.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/achievement_highway.php?organize=all" class="link-btn">全局</a>
                </div>
                <div class="card">
                    <h3>高快速公路</h3>
                    <a href="./report/highway.php?organize=one" class="link-btn">一分局</a>
                    <a href="./report/highway.php?organize=two" class="link-btn">二分局</a>
                    <a href="./report/highway.php?organize=three" class="link-btn">三分局</a>
                    <a href="./report/highway.php?organize=traffic" class="link-btn">交通隊</a>
                    <a href="./report/highway.php?organize=all" class="link-btn">全局</a>
                </div>
            </div>
            <div class="grid-bottom">
                <div class="card">
                    <h3>整理交通秩序(惡性違規)</h3>
                    <div class="card-links-2col">
                        <a href="./report/statistics_record.php?period=current" class="link-btn">月報-當期</a>
                        <a href="./report/statistics_record.php?period=total" class="link-btn">月報-累計</a>
                        <a href="./report/statistics_public.php?period=current" class="link-btn">民眾檢舉-當期</a>
                        <a href="./report/statistics_public.php?period=total" class="link-btn">民眾檢舉-累計</a>
                        <a href="./report/statistics_technology_law.php?period=current" class="link-btn">科技執法-當期</a>
                        <a href="./report/statistics_technology_law.php?period=total" class="link-btn">科技執法-累計</a>
                        <a href="./report/statistics_remove.php?period=current" class="link-btn">剔除檢舉-當期</a>
                        <a href="./report/statistics_remove.php?period=total" class="link-btn">剔除檢舉-累計</a>
                        <a href="./report/statistics_police.php?period=current" class="link-btn">員警舉發-當期</a>
                        <a href="./report/statistics_police.php?period=total" class="link-btn">員警舉發-累計</a>
                    </div>
                </div>
                <div class="card">
                    <h3>其他</h3>
                    <div class="card-links-2col">
                        <a href="./report/phone_table_inventory.php" class="link-btn">手持使用行動電話統計表(清冊)</a>
                        <a href="./report/phone_table.php" class="link-btn">手持使用行動電話統計表</a>
                        <a href="./report/dangerous_driving.php" class="link-btn">防制危險駕車績效成果表</a>
                        <a href="./report/drunk_driving.php" class="link-btn">取締酒駕</a>
                        <a href="./report/manual_driving.php" class="link-btn">建檔-手開單</a>
                        <a href="./report/order_revocation.php" class="link-btn">建檔-電腦撤掣單</a>
                        <a href="./report/dangerous_driving_sample.php" class="link-btn">R12防制危險駕車績效成果表上傳範本</a>
                        <a href="./report/double_parking.php" class="link-btn">並排停車舉發統計</a>
                        <a href="./report/43_law.php" class="link-btn">取締43條</a>
                        <a href="./report/large_car.php" class="link-btn">取締大型車</a>
                        <a href="./report/technology_law.php" class="link-btn">取締科技執法</a>
                        <a href="./report/exhaust_pipe.php" class="link-btn">取締排氣管</a>
                        <a href="./report/plate_clean_up.php" class="link-btn">取締淨牌</a>
                        <a href="./report/no_yielding_to_pedestrians.php" class="link-btn">取締路口不停讓行人</a>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script></script>
</body>

</html>