<!doctype html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>案件查詢系統</title>

    <style></style>
</head>

<body>
    <?php include 'header.php'; ?>

    <div class="container-fluid">
        <div class="main-container">
            <?php
            $breadcrumbs = [
                ['label' => '首頁', 'url' => 'action_page.php', 'color' => 'text-secondary'],
                ['label' => '查詢日誌', 'url' => '', 'color' => 'text-dark']
            ];
            include 'breadcrumb.php';
            ?>

            <!-- 查詢條件 -->
            <div class="card p-4 mb-4">
                <form class="index-form" action="api/logs/export_log.php" accept-charset="UTF-8" method="post">
                    <div class="row">
                        <div class="col-md-3 col-sm-12 mb-2">
                            <label class="mb-1">查詢日期：</label>
                            <input type="date" class="form-control" name="query_date" id="query_date">
                        </div>
                    </div>

                    <div class="d-flex justify-end mt-3">
                        <button type="submit" class="btn btn-primary w-100 w-md-auto" id="search_btn">匯出</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script>
        $(document).ready(function() {
            const today = new Date().toISOString().split('T')[0];
            $('#query_date').val(today);
        })
    </script>
</body>

</html>