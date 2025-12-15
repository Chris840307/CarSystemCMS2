<?php
/**
 * breadcrumb.php
 * 
 * 使用範例：
 * $breadcrumbs = [
 *   ['label' => '首頁', 'url' => 'action_page.php', 'color' => 'text-primary'],
 *   ['label' => '監理入案作業', 'url' => '', 'color' => 'text-secondary']
 * ];
 */

if (!isset($breadcrumbs) || !is_array($breadcrumbs)) return;

?>

<nav aria-label="breadcrumb" class="my-2">
    <ol class="breadcrumb align-items-center mb-0 bg-transparent p-0">
        <?php foreach ($breadcrumbs as $index => $crumb): ?>
            <li class="breadcrumb-item d-flex align-items-center <?php echo ($index === count($breadcrumbs)-1) ? 'active' : ''; ?>" 
                <?php echo ($index === count($breadcrumbs)-1) ? 'aria-current="page"' : ''; ?>>

                <?php if (!empty($crumb['url']) && $index !== count($breadcrumbs)-1): ?>
                    <a href="<?php echo $crumb['url']; ?>" 
                       class="text-decoration-none <?php echo $crumb['color'] ?? 'text-primary'; ?> fw-bold d-flex align-items-center">
                        <?php echo $crumb['label']; ?>
                    </a>
                <?php else: ?>
                    <span class="<?php echo $crumb['color'] ?? 'text-secondary'; ?> fw-bold"><?php echo $crumb['label']; ?></span>
                <?php endif; ?>
            </li>

            <?php if ($index !== count($breadcrumbs)-1): ?>
                <!-- 分隔符號 -->
                <span class="mx-1 d-flex align-items-center">
                    <img src="./assets/svg/ChevronRight.svg" width="16" height="16" alt=">" />
                </span>
            <?php endif; ?>
        <?php endforeach; ?>
    </ol>
</nav>
