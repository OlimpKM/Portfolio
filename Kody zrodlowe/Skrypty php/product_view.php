<?php

/**
 * Product Display Engine
 * Handles database interaction, template rendering, and product categorization.
 */

// Load core framework libraries
$lib_path = 'framework/core.inc.php';
if (!file_exists($lib_path)) { die("Critical Error: Core library missing."); }
require_once($lib_path);

// Initialize system paths
$path_framework = create_path(array('framework'));
$path_database  = create_path(array('vendor', 'adodb'));
$path_templates = create_path(array('templates', 'product_engine'));

// Load required modules
require_once($path_framework . 'error_handler.inc.php');
require_once($path_framework . 'template_engine.inc.php');
require_once($path_database  . 'adodb.inc.php');
require_once($path_framework . 'html_utils.inc.php');
require_once($path_framework . 'registry.inc.php');

// Error reporting configuration
error_reporting(E_ERROR | E_WARNING | E_PARSE);
set_error_log('system_errors.log');

// Initialize database connection
$db = &ADONewConnection('mysqli');
if (!$db->Connect($db_host, $db_user, $db_pass, $db_name)) {
    handle_system_error("Database connection failed.");
}

// Prepare request parameters
$action     = isset($_GET['action']) ? $_GET['action'] : 'list';
$category_id = isset($_GET['cat_id']) ? (int)$_GET['cat_id'] : 0;
$product_id  = isset($_GET['prod_id']) ? (int)$_GET['prod_id'] : 0;

// Initialize View Variables
$content_body = "";
$page_title   = "Product Catalog";

try {
    // Logic for fetching categories
    $sql_categories = "SELECT id, name FROM categories WHERE active = 1 ORDER BY sort_order";
    $rs_categories = $db->Execute($sql_categories);

    $category_list = "";
    if ($rs_categories && !$rs_categories->EOF) {
        while (!$rs_categories->EOF) {
            $cat_name = $rs_categories->fields['name'];
            $cat_link = "product_view.php?cat_id=" . $rs_categories->fields['id'];
            $category_list .= "<li><a href=\"$cat_link\">$cat_name</a></li>";
            $rs_categories->MoveNext();
        }
    }

    // Logic for fetching products based on selection
    if ($category_id > 0) {
        $sql_products = "SELECT * FROM products WHERE category_id = " . $db->qstr($category_id);
        $rs_products = $db->Execute($sql_products);
        // ... processing product list ...
    }

} catch (Exception $e) {
    $error_msg = "An error occurred while processing your request.";
    header("Location: error_page.php?msg=" . urlencode($error_msg));
    exit;
}

// Finalize database connection
$db->Close();

// --- Template Rendering ---

$view = new Template($path_templates . 'product_body.tpl');
$view->add('app_name', 'Global Product Catalog');
$view->add('category_html', $category_list);
$view->add('product_details', $product_data);
$view->add('info_message', $status_msg);

$page_body = $view->execute();

// Construct final HTML document
$layout = new Template($main_layout_template);
$layout->add('title', $page_title);
$layout->add('content', $page_body);
$layout->add('footer_text', '© ' . date('Y') . ' Product Engine. All rights reserved.');

echo $layout->execute();

?>