<!doctype html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>案件查詢系統</title>

    <style>
        .card {
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        }

        .section-title {
            font-size: 20px;
            color: #1976d2;
            margin-bottom: 15px;
            font-weight: 600;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        table td,
        table th {
            padding: 6px 10px;
            border: 1px solid #ccc;
            background: white;
        }

        table th {
            width: 160px;
            background: #e9f1ff;
            border: 1px solid #b5cff4;
            white-space: nowrap;
        }

        /* 手機版 RWD：table 變成卡片 */
        @media (max-width: 768px) {

            table,
            tbody,
            tr,
            td,
            th {
                display: block;
                width: 100%;
            }

            tr {
                margin-bottom: 12px;
                border: 1px solid #ddd;
                padding: 10px;
                border-radius: 6px;
                background: #fff;
            }

            td,
            th {
                border: none !important;
                padding: 4px 0 !important;
            }

            th {
                font-weight: 700;
                color: #1a4e8a;
                margin-top: 6px;
            }

            td {
                margin-bottom: 6px;
            }

            /* 隱藏多餘分隔線 */
            table.table-bordered {
                border: none !important;
            }
        }

        .image-box {
            border: 1px solid #ddd;
            padding: 10px;
            background: #fff;
            text-align: center;
        }

        .image-box img {
            max-width: 100%;
            border-radius: 4px;
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
                ['label' => '舊資料查詢', 'url' => 'case_old.php', 'color' => 'text-dark'],
                ['label' => '資料明細', 'url' => '', 'color' => 'text-dark']
            ];
            include 'breadcrumb.php';
            ?>
            <div class="card p-4">
                <div class="section-title">舊案件資料明細</div>
                <div class="row">
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">舉發單位</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="department" id="department" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">舉發員警</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="officer" id="officer" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">違規單號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="form_No" id="form_No" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">違規車號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="plate_number" id="plate_number" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">車別</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="vehicle_type" id="vehicle_type" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">入案日期</span>
                        <div class="input-wrapper w-100p">
                            <input type="date" class="form-control" name="form_date" id="form_date" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">違規日期</span>
                        <div class="input-wrapper w-100p">
                            <input type="date" class="form-control" name="violation_date" id="violation_date" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">違規人</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="name" id="name" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-6">
                        <span class="form-label w-100">違反法條1</span>
                        <div class="input-wrapper w-100p">
                            <select class="form-control" name="law1" id="law1" disabled> <!-- append option--> </select>
                        </div>
                    </div>
                    <div class="form-group col-md-6">
                        <span class="form-label w-100">違反法條2</span>
                        <div class="input-wrapper w-100p">
                            <select class="form-control" name="law2" id="law2" disabled> <!-- append option--> </select>
                        </div>
                    </div>
                    <div class="form-group col-md-6">
                        <span class="form-label w-100">違規地點</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="violation_location" id="violation_location" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">違規人證號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="national_No" id="national_No" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">違規人地址</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="addr" id="addr" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">車主</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="car_owner" id="car_owner" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">車主證號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="car_No" id="car_No" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">車主地址</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="car_addr" id="car_addr" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">送達註記</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="delivery_note" id="delivery_note" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">到案日期</span>
                        <div class="input-wrapper w-100p">
                            <input type="date" class="form-control" name="arrive_date" id="arrive_date" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">裁決所屬</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="station" id="station" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">扣押內容</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="content" id="content" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">白單單號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="whiteNo" id="whiteNo" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">一次銷案</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="close_case1" id="close_case1" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">大宗掛號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="bulk1" id="bulk1" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">二次銷案</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="close_case2" id="close_case2" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">大宗掛號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="bulk2" id="bulk2" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">三次銷案</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="close_case3" id="close_case3" disabled>
                        </div>
                    </div>
                    <div class="form-group col-md-3">
                        <span class="form-label w-100">大宗掛號</span>
                        <div class="input-wrapper w-100p">
                            <input type="text" class="form-control" name="bulk3" id="bulk3" disabled>
                        </div>
                    </div>
                </div>

                <hr>

                <div class="row">
                    <div class="col-md-6 col-sm-12 mb-3">
                        <h5>違規影像</h5>
                        <div class="image-box">
                            <img id="image1" src="no-image.png" alt="影像不存在">
                        </div>
                    </div>

                    <div class="col-md-6 col-sm-12 mb-3">
                        <h5>送達證書影像</h5>
                        <div class="image-box">
                            <img id="image2" src="no-image.png" alt="影像不存在">
                        </div>
                    </div>
                </div>
                <!-- 功能 -->
                <div class="d-flex justify-center text-center mt-auto w-100p gap-2">
                    <button class="btn btn-outline-primary" id="prev_btn">上一頁</button>
                    <button class="btn btn-success" id="export_btn">匯出 Excel</button>
                </div>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script>
        $(document).ready(function() {
            // 取得明細資料
            fetchCaseDetail();
        })

        // 取得明細資料
        function fetchCaseDetail() {
            let postData = {
                id: getQueryParam('id')
            };
            $.ajax({
                url: 'api/violation/get_detail.php',
                type: 'POST',
                dataType: 'json',
                contentType: 'application/json',
                data: JSON.stringify(postData),
                success: function(res) {
                    if (res.returnCode == 200) {
                        $('#department').val(res.data.department);
                        $('#officer').val(res.data.officer);
                        $('#form_No').val(res.data.form_No);
                        $('#plate_number').val(res.data.plate_number);
                        $('#vehicle_type').val(res.data.vehicle_type);
                        $('#form_date').val(res.data.form_date);
                        $('#violation_date').val(res.data.violation_date);
                        $('#name').val(res.data.name);
                        $('#law1').val(res.data.law1);
                        $('#law2').val(res.data.law2);
                        $('#violation_location').val(res.data.location);
                        $('#national_No').val(res.data.national_No);
                        $('#addr').val(res.data.addr);
                        $('#car_owner').val(res.data.car_owner);
                        $('#car_No').val(res.data.car_No);
                        $('#car_addr').val(res.data.car_addr);
                        $('#delivery_note').val(res.data.delivery_note);
                        $('#arrive_date').val(res.data.arrive_date);
                        $('#station').val(res.data.station);
                        $('#content').val(res.data.content);
                        $('#whiteNo').val(res.data.whiteNo);
                        $('#close_case1').val(res.data.close_case1);
                        $('#bulk1').val(res.data.bulk1);
                        $('#close_case2').val(res.data.close_case2);
                        $('#bulk2').val(res.data.bulk2);
                        $('#close_case3').val(res.data.close_case3);
                        $('#bulk3').val(res.data.bulk3);
                        $("#image1").attr("src", d.image1);
                        $("#image2").attr("src", d.image2);
                    }
                },
                error: function(xhr) {
                    console.error("錯誤：", xhr.responseJSON.message);
                }
            });
        };

        // 上一頁
        $('#prev_btn').click(function() {
            window.location.replace('case_olres.data.php');
        });

        // 匯出
        $('#export_btn').click(function() {
            alert('匯出 Excel 功能尚未完成');
        });
    </script>
</body>

</html>