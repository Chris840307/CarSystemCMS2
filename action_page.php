<!doctype html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>案件查詢系統</title>

    <style>
        .container {
            max-width: 1200px;
            min-height: calc(100vh - 206px);
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

        #card_list {
            display: grid;
            gap: 12px;
        }
    </style>
</head>

<body>
    <?php include 'header.php'; ?>

    <div class="container">
        <!-- 功能列表 -->
        <h2 class="section-title">功能列表</h2>
        <div class="grid">

        </div>
        <div class="grid">
            <div class="card">
                <h3>案件管理</h3>
                <div id="card_list"></div>
                <!-- append data -->
            </div>
            <div class="card">
                <h3>統計報表</h3>
                <a href="./report_general.php" class="link-btn">一般報表</a>
            </div>
        </div>
    </div>

    <?php include 'footer.php'; ?>

    <script>
        let myPermissions = []; // 保存使用者權限名稱

        $(document).ready(function() {
            fetchPermissionList();
        })

        // 取得使用者權限
        function fetchPermissionList() {
            const postData = {
                user_id: getCookie('user_id'),
            };

            $.ajax({
                url: 'api/acl_users/get_user_list.php',
                type: 'POST',
                dataType: 'json',
                contentType: 'application/json',
                data: JSON.stringify(postData),
                success: function(res) {
                    if (res.returnCode == 200) {
                        myPermissions = res.data[0].permissions.map(p => p.name);
                        renderDashboard();
                    }
                },
                error: function(xhr) {
                    console.error("錯誤：", xhr.responseJSON.message);
                }
            });
        };

        function renderDashboard() {
            const container = document.getElementById('card_list');
            container.innerHTML = ''; // 先清空

            // 權限對應按鈕資料
            const links = [{
                    permission: '宏謙案件查詢',
                    href: './case.php',
                    text: '宏謙案件查詢'
                },
                {
                    permission: '舊案件查詢',
                    href: './case_old.php',
                    text: '舊案件查詢'
                },
                {
                    permission: '查詢日誌',
                    href: './logs.php',
                    text: '查詢日誌'
                },
                {
                    permission: '權限管理',
                    href: './acl_users.php',
                    text: '權限管理'
                },
            ];

            // 根據權限生成按鈕
            links.forEach(link => {
                if (myPermissions.includes(link.permission)) {
                    const a = document.createElement('a');
                    a.href = link.href;
                    a.className = 'link-btn';
                    a.textContent = link.text;
                    container.appendChild(a);
                }
            });
        }
    </script>
</body>

</html>