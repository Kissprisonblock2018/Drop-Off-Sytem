<?php
session_start();
require_once 'db.php';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $shopName = trim($_POST['shop_name']);
    $country = $_POST['country'];
    $state = $_POST['state'];
    $postal = $_POST['postal'];
    $addr1 = $_POST['address1'];
    $addr2 = $_POST['address2'];
    $city = $_POST['city'];
    $phone = $_POST['phone'];
    $aspect = $_POST['aspect'] ?? null;
    $logoPath = null;

    if (!empty($_FILES['logo']['name'])) {
        $targetDir = 'uploads/';
        if (!is_dir($targetDir)) {
            mkdir($targetDir, 0777, true);
        }
        $logoPath = $targetDir . basename($_FILES['logo']['name']);
        move_uploaded_file($_FILES['logo']['tmp_name'], $logoPath);
    }

    $stmt = $pdo->prepare('INSERT INTO seller_onboarding
        (shop_name, country, state, postal_code, address1, address2, city, phone, logo_path, aspect_ratio, progress)
        VALUES (?,?,?,?,?,?,?,?,?,?,25)');
    $stmt->execute([$shopName, $country, $state, $postal, $addr1, $addr2, $city, $phone, $logoPath, $aspect]);
    $_SESSION['seller_id'] = $pdo->lastInsertId();
    header('Location: step2.php');
    exit();
}
?>
<!DOCTYPE html>
<html>
<head>
    <title>Shop Information</title>
    <link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>
</head>
<body>
<h1>Shop Information</h1>
<div id="progress"></div>
<form method="post" enctype="multipart/form-data">
    <label>Shop Name*<br><input type="text" name="shop_name" required></label><br>
    <label>Country*<br><input type="text" name="country" required></label><br>
    <label>State*<br><input type="text" name="state" required></label><br>
    <label>Postal Code*<br><input type="text" name="postal" required></label><br>
    <label>Address Line 1*<br><input type="text" name="address1" required></label><br>
    <label>Address Line 2<br><input type="text" name="address2"></label><br>
    <label>City*<br><input type="text" name="city" required></label><br>
    <label>Phone*<br><input type="text" name="phone" required></label><br>
    <label>Aspect Ratio<br>
        <select name="aspect">
            <option value="1:1">1:1</option>
            <option value="free">Free</option>
        </select>
    </label><br>
    <label>Logo<br><input type="file" name="logo"></label><br>
    <button type="submit">Next Step</button>
</form>
<script>
$(function(){
    $("#progress").progressbar({value:25});
});
</script>
</body>
</html>
