<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Retrieve form data and sanitize inputs
    $name = htmlspecialchars($_POST["name"]);
    $email = htmlspecialchars($_POST["email"]);
    $phone = htmlspecialchars($_POST["phone"]);

    // Define the file where the data will be stored
    $file = "submissions.txt";

    // Format the data to be saved
    $data = "Name: $name\nEmail: $email\nPhone: $phone\n----------------------\n";

    // Save the data into the text file
    file_put_contents($file, $data, FILE_APPEND);

    // Redirect to a "thank you" page
    header("Location: thank-you.html");
    exit();
}
?>
