 <!-- #include virtual="/Hmath/func/DB.fun"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title></title>
<!--<base target="main">-->
<base target="_blank">
</head>

<body background="../hmathw/bg%20flowers1.bmp">
 
 
 <% 
   
 MySelf = Request.ServerVariables("PATH_INFO")

 ITEM=Request("ITEM")
 SECTM=Request("SECTM")
  'Dim URL
  'URL = SECTM&ITEM
 'URL = SECTM&"login-1.asp" 
 If Request("Send") <> Empty Then

     ' Response.Redirect "Test.asp?" & Request.QueryString
     'Response.Redirect "TestKa.asp?" & Request.QueryString
   Response.Redirect "login-1r.asp?" & Request.QueryString
   
    ' Response.Redirect URL 
   ''' window.open URL 
   '' <option value="Dnt1/Dntn.htm">閱讀課本</option>
   ''  <INPUT Type=Hidden Name=SECTM Value="Dnt1/Dntn.htm""DNT"&Trim(Right(<?=SECTM?>,1))&"/Dntnh.htm">    
   ''  <option value="Dnt1/Dntn.htm">閱讀課本</option>
   
 End If     
   
  %>    

<h2 align="left">數學天地</h2>
<h4>請選年級:& 項目:</h4>                 
                                                                                             
<blockquote>
  <form action="<%=Myself%>" method="GET">
    <p><select name="SECTM" size="2">                        
        <option value="http://61.222.248.199/Hmath-1/">高一(上)</option> 
        <option value="http://61.222.248.199/Hmath-2/">高一(下)</option> 
        <option value="http://61.222.248.199/Hmath-3/">高二(上)</option> 
        <option value="http://61.222.248.199/Hmath-4/">高二(下)</option>
    <!--<option value="http://61.222.248.199/TBKFCSH/TBKFX.htm">歷屆考題(複習教材)</option>--> 
    <!-- URL="http://localhost/Hsu-fundb/enterr2-1.asp" -->                
   </select>                   
       </p>  
        
    <p><select name="ITEM" size="3">                                        
        <option value="TBKIN/indxbook.htm">閱讀課本</option>
        <option value="Enterkac-1t.asp">逐題練習</option> 
        <option value="Enterkac-1.asp">平常測驗</option>
       </select> 
                                                                                                                                                                                                                                      
      <p><input type="submit" Name="Send" value=" 進入考場 "> </p>
   </form> 
  </blockquote> 
  <h2 align="left"><a href="http://61.222.248.199/TBKFCSH/TBKFX.htm" >複習教材</a></h2>
<!--<h2 align="left"><a href="http://61.222.248.199/hsu-fundb/login-2r.asp" >紫微斗數</a></h2>-->
  <h2 align="left"><a href="http://61.222.248.199/hsu-fundb/enterr2-1.asp" >紫微斗數</a></h2>
  <h2 align="left"><a href="http://61.222.248.199/hsu-pk/enterr2-pk.asp" >易經卜卦</a></h2>
   <h2 align="left"><a href="http://ms.ho.net.tw/" >休閒園地</a></h2>
<!--<h2 align="left"><a href="http://61.222.248.199/Hmath-1ds/main.asp" >購物街</a></h2>-->
<!-- <h2 align="left"><a href="http://class.ruten.com.tw/user/index00.php?s=tang1206" >拍賣場</a></h2>-->
 <h2 align="left"><a href="http://61.222.248.199/Hmath/Rm-ch11/gform.htm" >留言版</a></h2>
 
<!-- <h2 align="left"><a href="http://promotion.ezpay.com.tw/c2c/mentoring/what_ezpay/chart02.htm"><img src="http://promotion.ezpay.com.tw/c2c/logo/icon9.gif" border="0" alt="ezPay 個人帳房"></a></h2>-->
<h2 align="left"><a href="https://www.ezpay.com.tw/ezPay_C2C/welcome/index.jsp"><img src="http://promotion.ezpay.com.tw/c2c/logo/icon9.gif" border="0" alt="ezPay 個人帳房"></a></h2>
 
<form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_s-xclick">
<input type="hidden" name="hosted_button_id" value="8QCVEU3W7PRUA">
<input type="image" src="https://www.paypalobjects.com/zh_TW/TW/i/btn/btn_cart_LG.gif" border="0" name="submit" alt="PayPal － 更安全、更簡單的線上付款方式！">
<img alt="" border="0" src="https://www.paypalobjects.com/zh_TW/i/scr/pixel.gif" width="1" height="1">
</form>

  
 <hr> 
<FONT Color=Red><%=Msg%></FONT> 
</boody> 
 
 
</html>






















































