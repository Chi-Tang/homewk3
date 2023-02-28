
<!DOCTYPE html>
<html lang="zh-cn">
<head>
<meta charset="utf-8" />
<title>writing-mode_CSS参考手册_web前端开发参考手册系列</title>
<meta name="author" content="Joy Du(飘零雾雨), dooyoe@gmail.com, www.doyoe.com" />
<style type ="text/css">
.test{width:100px;height:100px;margin:10px;border:1px solid #aaa;}
.lr-tb{-ms-writing-mode:lr-tb;}
.tb-rl{-ms-writing-mode:tb-rl;}
.tb-rl{-webkit-writing-mode:vertical-rl;}
</style>

</head>
<body>
 <div style="-ms-writing-mode: tb-rl">Content rendered vertically</div>
<div class="test lr-tb">本段文字将按照水平从左到右的书写方向进行流动。</div>
<div class="test tb-rl">本段文字将按照垂直从右到左的书写方向进行流动。</div>
<ul class="ttb-rl">
	<li class="test">本段文字将按照垂直从右到左的书写方向进行流动。</li>
	<li class="test">本段文字将按照垂直从右到左的书写方向进行流动。</li>
</ul>

</body>
</html>










