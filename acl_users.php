<!doctype html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>案件查詢系統</title>

    <style>
        #result_table tbody tr:hover td {
            background-color: #dce8ff;
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
                ['label' => '權限管理', 'url' => '', 'color' => 'text-dark']
            ];
            include 'breadcrumb.php';
            ?>

            <div class="card p-4 mb-4">
                <table class="table table-custom" id="result_table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>帳號</th>
                            <th>姓名</th>
                            <th>單位</th>
                            <th>權限</th>
                        </tr>
                    </thead>
                    <tbody>
                        <!-- append data -->
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script>
        let currentData = []; // 保存查詢結果
        let allPermissions = []; // 保存權限結果

        $(document).ready(function() {
            Promise.all([
                    fetchPermissionList(),
                    fetchUserList()
                ])
                .then(() => {
                    renderTable();
                })
                .catch(err => {
                    console.error('載入資料失敗', err);
                    alert('資料載入失敗');
                });
        })

        // 取得帳號清單
        function fetchUserList() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'api/acl_users/get_user_list.php',
                    type: 'GET',
                    dataType: 'json',
                    contentType: 'application/json',
                    success: function(res) {
                        if (res.returnCode == 200) {
                            currentData = res.data;
                            resolve(res.data);
                        } else {
                            reject(res);
                        }
                    },
                    error: function(xhr) {
                        console.error("錯誤：", xhr.responseJSON.message);
                    }
                });
            });
        };
        // 取得權限清單
        function fetchPermissionList() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'api/acl_users/get_permission_list.php',
                    type: 'GET',
                    dataType: 'json',
                    contentType: 'application/json',
                    success: function(res) {
                        if (res.returnCode == 200) {
                            allPermissions = res.data;
                            resolve(res.data);
                        } else {
                            reject(res);
                        }
                    },
                    error: function(xhr) {
                        console.error("錯誤：", xhr.responseJSON.message);
                    }
                });
            });
        };

        // 渲染表格
        function renderTable() {
            const tbody = $("#result_table tbody");
            tbody.empty();

            if (currentData.length === 0) {
                tbody.append(`<tr><td colspan="6" class="text-center">查無資料</td></tr>`);
                return;
            }

            currentData.forEach(row => {
                // 組權限 checkbox
                let permissionHtml = '';
                allPermissions.forEach(p => {
                    const checked = row.permissions.some(up => up.id === p.id) ? 'checked' : '';

                    permissionHtml += `
                        <label style="margin-right:12px; white-space:nowrap;">
                            <input type="checkbox"
                                data-user-id="${row.id}"
                                data-permission-id="${p.id}"
                                ${checked}
                                onchange="togglePermission(this)">
                            ${p.name}
                        </label>
                    `;
                });

                tbody.append(`
                        <tr>
                            <td>${row.id}</td>
                            <td>${row.acc}</td>
                            <td>${row.name}</td>
                            <td>${row.department}</td>
                            <td>${permissionHtml}</td>
                        </tr>
                    `);
            });
        }

        // 更新權限
        function togglePermission(el) {
            const postData = {
                user_id: $(el).data('user-id'),
                permission_id: $(el).data('permission-id'),
                value: el.checked
            };

            $.ajax({
                url: 'api/acl_users/update_user_permission.php',
                method: 'POST',
                contentType: 'application/json',
                data: JSON.stringify(postData),
                success(res) {
                    if (res.returnCode == 200) {
                        showToast('更新成功', 'success');
                    } else {
                        showToast('更新失敗', 'error');
                        el.checked = !el.checked;
                    }
                },
                error: function(xhr) {
                    showToast('更新失敗', 'error');
                    console.error("錯誤：", xhr.responseJSON.message);
                    el.checked = !el.checked;
                }
            });
        }
    </script>
</body>

</html>