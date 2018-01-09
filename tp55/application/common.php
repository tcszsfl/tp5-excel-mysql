<?php
// +----------------------------------------------------------------------
// | ThinkPHP [ WE CAN DO IT JUST THINK ]
// +----------------------------------------------------------------------
// | Copyright (c) 2006-2016 http://thinkphp.cn All rights reserved.
// +----------------------------------------------------------------------
// | Licensed ( http://www.apache.org/licenses/LICENSE-2.0 )
// +----------------------------------------------------------------------
// | Author: 流年 <liu21st@gmail.com>
// +----------------------------------------------------------------------

// 应用公共文件
use PHPMailer\PHPMailer\PHPMailer;
function sendmailer($tomail,$title,$body){
	$mail=new PHPMailer();
	$toemail=$tomail;//定义收件人的邮箱
	$mail->isSMTP();// 使用SMTP服务
	$mail->CharSet='UTF-8';
    $mail->SMTPAuth = true;// 启用SMTP验证
    $mail->Host = "smtp.126.com";// 发送方的SMTP服务器地址
    $mail->Username = 'tcszsfl@126.com';// 发送方的邮箱用户名，就是自己的邮箱名
    $mail->Password = 'sun1997';// 发送方的邮箱密码，不是登录密码,是qq的第三方授权登录码,要自己去开启,在邮箱的设置->账户->POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务 里面
    $mail->Port = 25;// QQ邮箱的ssl协议方式端口号是465/587
    $mail->setFrom('tcszsfl@126.com', '孙凤利'); //设置发件人   
     $mail->addAddress($toemail, 'sfl'); // 设置收件人信息，如邮件格式说明中的收件人
      $mail->addReplyTo('tcszsfl@126.com', 'Reply');// 设置回复人信息，指的是收件人收到邮件后，如果要回复，回复邮件将发送到的邮箱地址
      //$mail->addCC("xxx@163.com");// 设置邮件抄送人，可以只写地址，上述的设置也可以只写地址(这个人也能收到邮件)
    //$mail->addBCC("xxx@163.com");// 设置秘密抄送人(这个人也能收到邮件)
    //$mail->addAttachment("bug0.jpg");// 添加附件
     $mail->Subject = $title;
     $mail->Body    = $body;
     if(!$mail->send()){
     	  echo "Message could not be sent.";
        echo "Mailer Error: ".$mail->ErrorInfo;// 输出错误信息
     }else{
     	  echo '发送成功';
     }
}