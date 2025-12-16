<!doctype html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>案件查詢系統</title>

    <script src="js/jquery-3.7.1.min.js"></script>
    <script src="js/bootstrap-5.3.0.bundle.min.js"></script>
    <script src="js/jquery.validate-1.20.0.min.js"></script>
    <script src="js/messages_zh-1.20.0.min.js"></script>
    <script src="js/main.js?v20251217"></script>
    <script src="js/common.js?v20251115"></script>
    <link rel="stylesheet" href="css/bootstrap-5.3.0.min.css" />
    <link rel="stylesheet" href="css/style.css?v20251208" />
    <style>
        body {
            font-size: 16px;
            transition: font-size 0.2s ease;
        }

        .topbar {
            background: #1976d2;
            color: white;
            padding: 12px 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15);
        }

        .topbar-left {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .logo {
            width: 48px;
            height: 48px;
            border-radius: 50%;
            overflow: hidden;
            border: 2px solid white;
        }

        .logo img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .site-title {
            font-size: 20px;
            font-weight: bold;
        }

        .user-info {
            display: flex;
            align-items: center;
            gap: 16px;
            position: relative;
            font-size: 14px;
        }

        .nav-link.dropdown-toggle {
            padding-right: 15px;
            display: flex;
            align-items: center;
            gap: 6px;
            color: white;
        }

        .idle-timer {
            color: white;
            white-space: nowrap;
        }

        /* 懸浮字體控制按鈕 */
        .floating-font-controls {
            position: fixed;
            right: 20px;
            bottom: 20px;
            background: #1976d2;
            border-radius: 12px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.2);
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 8px;
            z-index: 999;
        }

        .font-btn {
            background: transparent;
            border: none;
            color: white;
            cursor: pointer;
            padding: 4px;
            padding-bottom: 0px;
            border-radius: 8px;
            transition: all 0.2s ease;
        }

        .font-btn:hover {
            background: rgba(255, 255, 255, 0.2);
        }

        .font-btn svg {
            width: 20px;
            height: 20px;
        }

        .font-label {
            color: white;
            font-size: 12px;
            margin-top: 4px;
            opacity: 0.85;
        }

        /* 錯誤訊息換行 */
        message-error {
            display: block;
            margin-top: 2px;
            color: red;
        }
    </style>
</head>

<body>
    <!-- 上方導覽列 -->
    <div class="topbar">
        <div class="topbar-left">
            <div class="logo">
                <img src="./assets/img/logo.jpg" alt="Logo">
            </div>
            <div class="site-title">宏謙實業有限公司案件查詢系統</div>
        </div>

        <div class="user-info">
            <!-- 使用者資訊 -->
            <ul class="navbar-nav ms-auto">
                <li class="nav-item dropdown no-arrow">
                    <a class="nav-link dropdown-toggle" href="#" id="userDropdown"
                        role="button" data-bs-toggle="dropdown" aria-expanded="false">
                        <svg width="20" height="20" fill="white" viewBox="0 0 24 24">
                            <path
                                d="M12 12c2.7 0 5-2.3 5-5s-2.3-5-5-5-5 
                                    2.3-5 5 2.3 5 5 5zm0 2c-3.3 0-10 
                                    1.7-10 5v3h20v-3c0-3.3-6.7-5-10-5z" />
                        </svg>
                        <span id="username"></span>
                    </a>
                    <ul class="dropdown-menu dropdown-menu-end" aria-labelledby="userDropdown">
                        <li><a class="dropdown-item" onclick="logout()">登出</a></li>
                    </ul>
                </li>
            </ul>

            <span class="idle-timer">
                <span id="idleTime">20:00</span> 後將自動登出
            </span>
        </div>
    </div>

    <!-- Custom Toast 容器 -->
    <div id="toast-container"></div>

    <!-- 懸浮字體控制區 -->
    <div class="floating-font-controls">
        <button class="font-btn" id="font-larger" title="放大字體">
            <svg fill="currentColor" viewBox="0 0 24 24">
                <path d="M19 11h-6V5h-2v6H5v2h6v6h2v-6h6z"></path>
            </svg>
        </button>
        <button class="font-btn" id="font-default" title="預設字體">
            <svg fill="currentColor" viewBox="0 0 24 24">
                <path d="M12 4l6 16h-2.25l-1.71-5H9.96l-1.71 5H6L12 4zm0 5.48L10.08 13h3.84L12 9.48z"></path>
            </svg>
        </button>
        <button class="font-btn" id="font-smaller" title="縮小字體">
            <svg fill="currentColor" viewBox="0 0 24 24">
                <path d="M19 13H5v-2h14v2z"></path>
            </svg>
        </button>
        <div class="font-label">字體</div>
    </div>

    <script>
        $(document).ready(function() {
            updateIdleDisplay();
            startIdleCountdown();

            cssControl() // 字體大小控制

            // 顯示使用者名稱
            let username = getCookie("name");
            if (username) {
                $("#username").text(username);
            } else {
                $("#username").text("未登入");
                // 若要未登入自動跳回登入頁 → 打開下面
                // window.location.href = "index.php";
            }
        });

        // 登出
        function logout() {
            // 清Cookie
            document.cookie = "user_id=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/";
            document.cookie = "name=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/";
            window.location.href = "index.php";
        }
    </script>
</body>

</html>