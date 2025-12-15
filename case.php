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
                ['label' => '宏謙案件查詢', 'url' => '', 'color' => 'text-dark']
            ];
            include 'breadcrumb.php';
            ?>

            <!-- 查詢條件 -->
            <div class="card p-4 mb-4">
                <div class="row">
                    <div class="col-md-3 col-sm-12 mb-2">
                        <label class="mb-1">違規單號</label>
                        <input type="text" class="form-control" name="form_No" id="form_No">
                    </div>
                    <div class="col-md-3 col-sm-12 mb-2">
                        <label class="mb-1">違規車號</label>
                        <input type="text" class="form-control" name="plate_number" id="plate_number">
                    </div>
                    <div class="col-md-3 col-sm-12 mb-2">
                        <label class="mb-1">違規日期</label>
                        <input type="date" class="form-control" name="violation_date" id="violation_date">
                    </div>
                    <div class="col-md-3 col-sm-12 mb-2">
                        <label class="mb-1">舉發單位</label>
                        <select class="form-control" name="department" id="department">
                            <option value="">請選擇</option>
                            <!-- append option-->
                        </select>
                    </div>
                    <div class="col-md-3 col-sm-12 mb-2">
                        <label class="mb-1">舉發員警</label>
                        <input type="text" class="form-control" name="officer" id="officer">
                    </div>
                </div>

                <div class="d-flex justify-end mt-3">
                    <button type="button" class="btn btn-outline-secondary mr-2 w-100 w-md-auto" id="clear_btn">清除</button>
                    <button type="button" class="btn btn-primary w-100 w-md-auto" id="search_btn">查詢</button>
                </div>
            </div>

            <!-- 查詢結果 -->
            <div class="card p-4 mb-4">
                <div class="section-title">查詢結果 共<span id="total"></span>筆</div>
                <!-- 待假資料倒入後改用API取得資料 -->
                <table class="table table-custom" id="result_table">
                    <thead>
                        <tr>
                            <th>違規單號</th>
                            <th>車號</th>
                            <th>違規日期</th>
                            <th>車別</th>
                            <th>案件狀態</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        <!-- append data -->
                    </tbody>
                </table>

                <div class="pagination">
                    <a href="#" class="page-btn prev">&lt;</a>
                    <a href="#" class="page-num active">1</a>
                    <a href="#" class="page-num">2</a>
                    <a href="#" class="page-num">3</a>
                    <a href="#" class="page-btn next">&gt;</a>
                </div>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script>
        let currentData = []; // 保存查詢結果
        let currentPage = 1; // 目前頁碼
        const rowsPerPage = 20; // 每頁筆數

        // 查詢
        $('#search_btn').click(function() {
            let postData = {
                form_No: $('#form_No').val(),
                plate_number: $('#plate_number').val(),
                violation_date: $('#violation_date').val(),
                department: $('#department').val(),
                officer: $('#officer').val()
            };
            $.ajax({
                url: 'api/case/get_list.php',
                type: 'POST',
                dataType: 'json',
                contentType: 'application/json',
                data: JSON.stringify(postData),
                success: function(res) {
                    if (res.returnCode === 200) {
                        currentData = res.data;
                        currentPage = 1;
                        renderTable();
                        renderPagination();
                    }
                }
            });
        });

        // 渲染表格
        function renderTable() {
            const tbody = $("#result_table tbody");
            tbody.empty();

            if (currentData.length === 0) {
                tbody.append(`<tr><td colspan="6" class="text-center">查無資料</td></tr>`);
                $('#total').text(0);
                return;
            }

            const start = (currentPage - 1) * rowsPerPage;
            const end = start + rowsPerPage;
            const pageData = currentData.slice(start, end);

            pageData.forEach(row => {
                const statusTag = getStatusTag(row.close_case1, row.close_case2, row.close_case3);
                tbody.append(`
                            <tr>
                                <td>${row.form_No}</td>
                                <td>${row.plate_number}</td>
                                <td>${row.violation_date}</td>
                                <td>${row.vehicle_type}</td>
                                <td>${statusTag}</td>
                                <td><a href="./case_detail.php?id=${row.id}" class="btn btn-sm btn-primary">檢視</a></td>
                            </tr>
                        `);
            });

            $('#total').text(currentData.length);
        }

        // 案件狀態
        function getStatusTag(c1, c2, c3) {
            if (c1 === '不舉發') return `<span class="tag tag-info">不舉發</span>`;
            if (c1 === '已舉發') return `<span class="tag tag-primary">已舉發</span>`;
            if (!c1 && !c2 && !c3) return `<span class="tag tag-warning">未處理</span>`;
            return `<span class="tag tag-default">其他</span>`;
        }

        // 渲染分頁
        function renderPagination() {
            const totalPages = Math.ceil(currentData.length / rowsPerPage);
            const pagination = $(".pagination");
            pagination.empty();

            if (totalPages <= 1) return; // 不需要分頁

            pagination.append(`<a href="#" class="page-btn prev">&lt;</a>`);

            for (let i = 1; i <= totalPages; i++) {
                pagination.append(`<a href="#" class="page-num ${i === currentPage ? 'active' : ''}">${i}</a>`);
            }

            pagination.append(`<a href="#" class="page-btn next">&gt;</a>`);
        }

        // 分頁
        $(document).on('click', '.page-num', function(e) {
            e.preventDefault();
            currentPage = parseInt($(this).text());
            renderTable();
            renderPagination();
        });

        $(document).on('click', '.next', function(e) {
            e.preventDefault();
            const totalPages = Math.ceil(currentData.length / rowsPerPage);
            if (currentPage < totalPages) {
                currentPage++;
                renderTable();
                renderPagination();
            }
        });

        $(document).on('click', '.prev', function(e) {
            e.preventDefault();
            if (currentPage > 1) {
                currentPage--;
                renderTable();
                renderPagination();
            }
        });

        // 清除
        $('#clear_btn').click(function() {
            $('#form_No').val('');
            $('#plate_number').val('');
            $('#violation_date').val('');
            $('#department').val('');
            $('#officer').val('');
            currentData = [];
            currentPage = 1;
            renderTable();
            renderPagination();
        });
    </script>
</body>

</html>