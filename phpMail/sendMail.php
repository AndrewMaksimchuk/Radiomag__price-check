<?php

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

require 'config.php';
require 'PHPMailer-master/src/Exception.php';
require 'PHPMailer-master/src/PHPMailer.php';
require 'PHPMailer-master/src/SMTP.php';

// Instantiation and passing `true` enables exceptions
$mail = new PHPMailer(true);

try {
    //Server settings
    $mail->SMTPDebug  = 2;
    $mail->isSMTP();
    $mail->Host       = $HOST;
    $mail->SMTPAuth   = $SMTPAUTH;
    $mail->Username   = $USERNAME;
    $mail->Password   = $PASSWORD;
    $mail->SMTPSecure = $SMTPSECURE;
    $mail->Port       = $PORT;
    $mail->CharSet    = "UTF-8";

    //Recipients
    $mail->setFrom('radiomagBot@radiomag.com.ua', 'radiomagBot');
    $mail->addAddress('stores@radiomag.com.ua', 'Всі магазини');               // Name is optional

    $mail->addAttachment('C:\Users\Maksymchuk\Desktop\pyPrice\New_price.xlsx');    // Optional name

    // Content
    $mail->isHTML(true);                                  // Set email format to HTML
    $mail->Subject = 'Зміни цін на товарах';
    $mail->Body    = 'Всім привіт!<br><br> До листа прикріплений файл з товарами у яких змінилася ціна. <br>Можете передивитися і оновити ціни на вітринах!<br><br>Це автоматичний лист!';
    $mail->AltBody = 'Всім привіт! До листа прикріплений файл з товарами у яких змінилася ціна. Можете передивитися і оновити ціни на вітринах!';

    $mail->send();
    echo 'Лист був відправлений!';
    echo("\n");
    echo("Кінець звязку!");
} catch (Exception $e) {
    echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
}
?>