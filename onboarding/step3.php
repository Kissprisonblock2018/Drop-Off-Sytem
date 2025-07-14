<?php
session_start();
require_once 'db.php';

if (!isset($_SESSION['seller_id'])) {
    header('Location: step1.php');
    exit();
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $shipping = $_POST['shipping'];
    $fixedCost = $_POST['fixed_cost'] ?? null;
    $freeThreshold = $_POST['free_threshold'] ?? null;
    $packageSize = $_POST['package_size'] ?? null;
    $pickup = isset($_POST['pickup']) ? 1 : 0;
    $returns = (int)($_POST['returns'] ?? 12);
    $cancellation = (int)($_POST['cancellation'] ?? 3);
    $callScheduled = isset($_POST['call_scheduled']) ? 1 : 0;

    $stmt = $pdo->prepare('UPDATE seller_onboarding SET shipping_option=?, fixed_rate_cost=?, free_ship_threshold=?, package_size=?, offer_pickup=?, return_days=?, cancel_days=?, call_scheduled=?, progress=75 WHERE id=?');
    $stmt->execute([$shipping, $fixedCost, $freeThreshold, $packageSize, $pickup, $returns, $cancellation, $callScheduled, $_SESSION['seller_id']]);
    header('Location: step4.php');
    exit();
}
?>
<!DOCTYPE html>
<html>
<head>
    <title>Shipping Configuration</title>
    <link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js"></script>
</head>
<body>
<h1>Shipping Configuration</h1>
<div id="progress"></div>
<form method="post">
    <label>Select Shipping Method*<br>
        <select name="shipping" id="shipping" required>
            <option value="free">Free Shipping</option>
            <option value="fixed">Fixed Rate</option>
            <option value="distance">Weight & Distance-Based</option>
            <option value="pickup">Pick-up Only</option>
            <option value="call">Book a Call</option>
        </select>
    </label><br>
    <div id="fixed_fields" style="display:none;">
        <label>Cost<br><input type="number" step="0.01" name="fixed_cost"></label><br>
    </div>
    <div id="distance_fields" style="display:none;">
        <label>Average Package Size<br>
            <select name="package_size">
                <option value="small">Small</option>
                <option value="medium">Medium</option>
                <option value="large">Large</option>
            </select>
        </label><br>
    </div>
    <div id="threshold_field" style="display:none;">
        <label>Free Shipping Threshold<br><input type="number" step="0.01" name="free_threshold"></label><br>
    </div>
    <label><input type="checkbox" name="pickup"> Offer In-store Pickup</label><br>
    <label>Return Eligibility Days<br><input type="number" name="returns" value="12"></label><br>
    <label>Cancellation Eligibility Days<br><input type="number" name="cancellation" value="3"></label><br>
    <div id="call_field" style="display:none;">
        <label><input type="checkbox" name="call_scheduled"> Calendly call scheduled</label><br>
    </div>
    <button type="submit">Next Step</button>
</form>
<script>
$(function(){
    $("#progress").progressbar({value:75});
    function toggleFields() {
        var val = $('#shipping').val();
        $('#fixed_fields, #distance_fields, #threshold_field, #call_field').hide();
        if(val === 'fixed') {
            $('#fixed_fields, #threshold_field').show();
        } else if(val === 'distance') {
            $('#distance_fields, #threshold_field').show();
        } else if(val === 'call') {
            $('#call_field').show();
        }
    }
    $('#shipping').on('change', toggleFields).trigger('change');
});
</script>
</body>
</html>
