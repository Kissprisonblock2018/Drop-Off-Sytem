<?php
session_start();
require_once 'db.php';

if (!isset($_SESSION['seller_id'])) {
    header('Location: step1.php');
    exit();
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $desc = $_POST['description'];
    $badges = isset($_POST['badges']) ? implode(',', $_POST['badges']) : null;
    $facebook = $_POST['facebook'] ?? null;
    $instagram = $_POST['instagram'] ?? null;
    $pinterest = $_POST['pinterest'] ?? null;
    $skip = isset($_POST['skip']) ? 1 : 0;
    $completed = 1;
    $date = date('Y-m-d H:i:s');

    $stmt = $pdo->prepare('UPDATE seller_onboarding SET description=?, badges=?, facebook=?, instagram=?, pinterest=?, skipped_optional=?, completed=?, submission_dt=?, progress=100 WHERE id=?');
    $stmt->execute([$desc, $badges, $facebook, $instagram, $pinterest, $skip, $completed, $date, $_SESSION['seller_id']]);
    echo "<h2>Onboarding Complete!</h2>";
    session_destroy();
    exit();
}
?>
<!DOCTYPE html>
<html>
<head>
    <title>Shop Details</title>
    <link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>
</head>
<body>
<h1>Shop Details</h1>
<div id="progress"></div>
<form method="post">
    <label>Shop Description<br><textarea name="description" cols="40" rows="4"></textarea></label><br>
    <label>Badges<br>
        <select name="badges[]" multiple size="5">
            <option value="Minority Owned">Minority Owned</option>
            <option value="Veteran Owned">Veteran Owned</option>
            <option value="Family Owned">Family Owned</option>
            <option value="B-Corp Certified">B-Corp Certified</option>
            <option value="Neurodivergent Owned">Neurodivergent Owned</option>
            <option value="LGBTQ+ Owned">LGBTQ+ Owned</option>
            <option value="Eco Smart">Eco Smart</option>
            <option value="In-Store">In-Store</option>
        </select>
    </label><br>
    <label>Facebook URL<br><input type="url" name="facebook"></label><br>
    <label>Instagram URL<br><input type="url" name="instagram"></label><br>
    <label>Pinterest URL<br><input type="url" name="pinterest"></label><br>
    <button type="submit" name="complete">Complete Setup</button>
    <button type="submit" name="skip">Skip and Finish</button>
</form>
<script>
$(function(){
    $("#progress").progressbar({value:100});
});
</script>
</body>
</html>
