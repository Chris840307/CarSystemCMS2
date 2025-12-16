<!DOCTYPE html>
<html lang="zh-Hant">

<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>案件查詢系統</title>
    <script src="js/jquery-3.7.1.min.js"></script>
    <script src="js/bootstrap-5.3.0.bundle.min.js"></script>
    <script src="js/jquery.validate-1.20.0.min.js"></script>
    <script src="js/messages_zh-1.20.0.min.js"></script>
    <link rel="stylesheet" href="css/bootstrap-5.3.0.min.css" />
    <link rel="stylesheet" href="css/style.css" />
    <style>
        html,
        body {
            height: 100%;
            margin: 0;
            font-family: 'Segoe UI', sans-serif;
        }

        /* 背景圖設定 */
        body {
            background: url("./assets/img/bg.png") no-repeat center center fixed;
            background-size: cover;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }

        /* 遮罩層讓文字更清晰 */
        .overlay {
            position: fixed;
            inset: 0;
            background: rgba(0, 0, 0, 0.003);
            z-index: -1;
        }

        /* 置中主體 */
        .main-container {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            position: relative;
            z-index: 1;
        }

        /* 系統標題 */
        /* 系統標題 */
        .system-title {
            text-align: center;
            margin-bottom: 25px;
        }

        .main-title {
            margin: 0;
            font-size: 1.7rem;
            font-weight: 700;
            color: #1e3c72;
            /* 深藍，穩重 */
            letter-spacing: 5px;
        }

        .second-title {
            margin: 10px 0 0;
            font-size: 1.5rem;
            font-weight: 500;
            color: #4a90e2;
            /* 較亮的藍色 */
            letter-spacing: 3px;
        }

        /* 登入卡片 */
        .login-box {
            width: 100%;
            max-width: 380px;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 123, 255, 0.25);
            background-color: white;
        }

        .login-box h2 {
            text-align: center;
            margin-bottom: 20px;
            color: #1e88e5;
        }

        .form-control:focus {
            border-color: #1e88e5;
            box-shadow: 0 0 5px rgba(30, 136, 229, 0.5);
        }

        button {
            width: 100%;
            background-color: #1e88e5;
            border: none;
        }

        button:hover {
            background-color: #1565c0;
        }

        .info-box {
            margin-top: 20px;
            font-size: 0.9rem;
            color: #333;
            border: 1px solid #1e88e5;
            border-radius: 8px;
            padding: 15px 20px;
            background-color: #e3f2fd;
            line-height: 1.5;
        }

        .info-box ul {
            padding-left: 1.2em;
        }

        .copyright {
            position: fixed;
            bottom: 8px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 0.85rem;
            color: #f0f0f0;
            text-shadow: 0 1px 4px rgba(0, 0, 0, 0.8);
            user-select: none;
        }
    </style>
</head>

<body>
    <!-- 背景半透明遮罩 -->
    <div class="overlay"></div>

    <!-- 登入主體 -->
    <div class="main-container">
        <div class="login-box">
            <div class="system-title">
                <div class="main-title">宏謙實業有限公司</div>
                <div class="second-title">案件查詢系統</div>
            </div>

            <form id="form_data" onsubmit="return false;" class="needs-validation" novalidate>
                <div class="form-group">
                    <label for="username">使用者帳號</label>
                    <input type="text" class="form-control" name="username" id="username" placeholder="請輸入帳號" required>
                </div>
                <div class="form-group">
                    <label for="pwd">密碼</label>
                    <input type="password" class="form-control" name="pwd" id="pwd" placeholder="請輸入密碼" required>
                </div>
                <button type="submit" class="btn btn-primary mt-4" onclick="login()">登入</button>
            </form>
            <div class="info-box">
                <strong>注意事項：</strong>
                <ul>
                    <li>密碼須包含至少8碼以上、同時具備英數字及特殊符號等</li>
                    <li>密碼錯誤三次鎖定該帳號15分鐘內不可再嘗試登入</li>
                    <li>登入後閒置20分鐘將自動登出</li>
                </ul>
            </div>
        </div>
    </div>

    <div class="copyright">COPYRIGHT © 2025</div>

    <script>
        $(document).ready(function() {
            // 驗證
            $('#form_data').validate({
                rules: {
                    username: {
                        required: true
                    },
                    pwd: {
                        required: true
                    },
                },
                errorClass: 'text-red-600 p3',
            });
        })

        function login() {
            if (!$('#form_data').valid()) {
                return;
            }

            let postData = {
                account: $("#username").val().trim(),
                pwd: $("#pwd").val().trim()
            };
            $.ajax({
                url: "api/login/login.php",
                type: "POST",
                dataType: 'json',
                contentType: 'application/json',
                data: JSON.stringify(postData),
                success: function(res) {
                    if (res.returnCode == 200) {
                        // 存 cookie 7天
                        document.cookie = "user_id=" + res.data.user_id + "; path=/; max-age=" + (7 * 24 * 60 * 60);
                        document.cookie = "name=" + res.data.name + "; path=/; max-age=" + (7 * 24 * 60 * 60);

                        setTimeout(() => {
                            location.replace("action_page.php");
                        }, 100);
                    } else {
                        alert('登入失敗');
                    }
                },
                error: function(xhr) {
                    alert(xhr?.responseJSON?.message);
                    console.error("錯誤：", xhr?.responseJSON?.message || xhr.responseText);
                }
            });
        }
    </script>
</body>

</html>