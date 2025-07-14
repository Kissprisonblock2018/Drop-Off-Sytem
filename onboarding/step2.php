<?php
session_start();
require_once 'db.php';

if (!isset($_SESSION['seller_id'])) {
    header('Location: step1.php');
    exit();
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $hasInventory = $_POST['has_inventory'];
    $filePath = null;
    $includeDrafts = isset($_POST['include_drafts']) ? 1 : 0;
    $madeToOrder = isset($_POST['made_to_order']) ? 1 : 0;
    $integration = isset($_POST['integration']) ? 1 : 0;

    if ($hasInventory === 'yes' && !empty($_FILES['inventory_file']['name'])) {
        $targetDir = 'uploads/';
        if (!is_dir($targetDir)) {
            mkdir($targetDir, 0777, true);
        }
        $filePath = $targetDir . basename($_FILES['inventory_file']['name']);
        move_uploaded_file($_FILES['inventory_file']['tmp_name'], $filePath);
    }

    $stmt = $pdo->prepare('UPDATE seller_onboarding SET has_inventory=?, inventory_file=?, include_drafts=?, made_to_order=?, integration_access=?, progress=50 WHERE id=?');
    $stmt->execute([$hasInventory === 'yes' ? 1 : 0, $filePath, $includeDrafts, $madeToOrder, $integration, $_SESSION['seller_id']]);
    header('Location: step3.php');
    exit();
}
?>
<!DOCTYPE html>
<html>
<head>
    <title>Inventory Management</title>
    <link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>
</head>
<body>
<h1>Inventory Management</h1>
<div id="progress"></div>
<form method="post" enctype="multipart/form-data">
    <label>Do you have existing inventory?*<br>
        <select name="has_inventory" id="has_inventory" required>
            <option value="yes">Yes</option>
            <option value="no">No</option>
        </select>
    </label><br>
    <div id="inventory_details" style="display:none;">
        <label>Upload Inventory File<br><input type="file" name="inventory_file"></label><br>
        <label><input type="checkbox" name="include_drafts"> Include drafts/inactive</label><br>
        <label><input type="checkbox" name="made_to_order"> Mark as made-to-order</label><br>
        <label><input type="checkbox" name="integration"> Enable integration access</label><br>
    </div>
    <button type="submit">Next Step</button>
</form>
<script>
$(function(){
    $("#progress").progressbar({value:50});
    $('#has_inventory').on('change', function(){
        if($(this).val() === 'yes') {
            $('#inventory_details').show();
        } else {
            $('#inventory_details').hide();
        }
    }).trigger('change');
});
</script>
</body>
</html>
