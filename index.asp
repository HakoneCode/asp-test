<%@ LANGUAGE="VBScript" CODEPAGE=65001 %>
<% Session.CodePage=65001 %>
<%
On Error Resume Next
Dim ErrMsg1
Dim ErrMsg2
Dim ErrMsg3
Dim ErrMsg4
Dim ErrMsg5,ErrMsg5_1,ErrMsg5_2,ErrMsg5_3,ErrMsg5_4
Dim ErrMsg6,ErrMsg6_1,ErrMsg6_2,ErrMsg6_3
Dim ErrMsg7,ErrMsg7_1,ErrMsg7_2
Dim ErrMsg8
Dim ErrMsg9
Dim intErrCheck
Dim strWords
Dim ErrFlag
Dim i,j
Dim comp


Set objBP = Server.CreateObject("basp21")

If Request.Form("SendFlag")=1 Then
'---　件名　---
	If Len(Request.Form("name"))= 0 Then
		ErrMsg1="<span class=""required"">「件名」を入力してください。<br /></span>"
	End If

'---　件名（カナ）　---
	If Len(Request.Form("nameKana"))= 0 Then
		ErrMsg2="<span class=""required"">「件名（カナ）」を入力してください。<br /></span>"
	End If


'---　送信元アドレス　---
	If Len(Request.Form("company"))= 0 Then
		ErrMsg3="<span class=""required"">「送信元アドレス」を入力してください。<br /></span>"
	End If


'---　所属名　---
	If Len(Request.Form("syozoku"))= 0 Then
		ErrMsg4="<span class=""required"">「所属名」を入力してください。<br /></span>"
	End If

'---　ご住所　---
	If Len(Request.Form("zip1"))= 0 Or Len(Request.Form("zip2"))= 0 Or Len(Request.Form("address1"))= 0 Or Len(Request.Form("address2"))= 0 Then
'		ErrMsg5 = "<span class=""required"">「ご住所」を入力してください。<br /></span>" & vbCrLf

		'---　郵便番号　---
		If Len(Request.Form("zip1"))= 0 Then
			ErrMsg5 = ErrMsg5 & "<span class=""required"">「郵便番号（前半）」を入力してください。<br /></span>" & vbCrLf
			ErrMsg5_1 = ErrMsg5_1 & "<span class=""required"">「郵便番号（前半）」を入力してください。<br /></span>" & vbCrLf
		End If

		If Len(Request.Form("zip2"))= 0 Then
			ErrMsg5 = ErrMsg5 & "<span class=""required"">「郵便番号（後半）」を入力してください。<br /></span>" & vbCrLf
			ErrMsg5_2 = ErrMsg5_2 & "<span class=""required"">「郵便番号（後半）」を入力してください。<br /></span>" & vbCrLf
		End If

		'---　都道府県　---
'		If Len(Request.Form("address1"))= 0 Then
'			ErrMsg5 = ErrMsg5 & "<span class=""required"">「都道府県」を選択してください。<br /></span>" & vbCrLf
'			ErrMsg5_3 = ErrMsg5_3 & "<span class=""required"">「都道府県」を選択してください。<br /></span>" & vbCrLf
'		End If

		'---　市区町村その他　---
		If Len(Request.Form("address2"))= 0 Then
			ErrMsg5 = ErrMsg5 & "<span class=""required"">「市区町村その他」を入力してください。<br /></span>" & vbCrLf
			ErrMsg5_4 = ErrMsg5_4 & "<span class=""required"">「市区町村その他」を入力してください。<br /></span>" & vbCrLf
		End If

	Else
		'---　郵便番号　---
		'前半
		intErrCheck = 0
		strWords = "1234567890"
		
		For i=1 To Len(Request.Form("zip1"))
			For j=1 To Len(strWords)
				If Mid(Request.Form("zip1"),i,1)=Mid(strWords,j,1) Then
					intErrCheck = intErrCheck + 1
				End If
			Next
		Next
		If intErrCheck <> Len(Request.Form("zip1")) Then
			ErrMsg5= ErrMsg5 & "<span class=""required"">「郵便番号（前半）」は半角数字で入力してください。<br /></span>"
			ErrMsg5_1= ErrMsg5_1 & "<span class=""required"">「郵便番号（前半）」は半角数字で入力してください。<br /></span>"
		Else
			If Len(Request.Form("zip1"))<> 3 Then
				ErrMsg5 = ErrMsg5 & "<span class=""required"">「郵便番号（前半）」は3桁の半角数字で入力してください。<br /></span>" & vbCrLf
				ErrMsg5_1 = ErrMsg5_1 & "<span class=""required"">「郵便番号（前半）」は3桁の半角数字で入力してください。<br /></span>" & vbCrLf
			End If
		End If

		'後半
		intErrCheck = 0

		For i=1 To Len(Request.Form("zip2"))
			For j=1 To Len(strWords)
				If Mid(Request.Form("zip2"),i,1)=Mid(strWords,j,1) Then
					intErrCheck = intErrCheck + 1
				End If
			Next
		Next
		If intErrCheck <> Len(Request.Form("zip2")) Then
			ErrMsg5= ErrMsg5 & "<span class=""required"">「郵便番号（後半）」は半角数字で入力してください。<br /></span>"
			ErrMsg5_21= ErrMsg5_2 & "<span class=""required"">「郵便番号（前半）」は半角数字で入力してください。<br /></span>"
		Else
			If Len(Request.Form("zip2"))<> 4 Then
				ErrMsg5 = ErrMsg5 & "<span class=""required"">「郵便番号（後半）」は4桁の半角数字で入力してください。<br /></span>" & vbCrLf
				ErrMsg5_2 = ErrMsg5_2 & "<span class=""required"">「郵便番号（後半）」は4桁の半角数字で入力してください。<br /></span>" & vbCrLf
			End If
		End If

	End If

'---　電話番号　---
	If Len(Request.Form("phone1"))= 0 Or Len(Request.Form("phone2"))= 0 Or Len(Request.Form("phone3"))= 0 Then
		ErrMsg6="<span class=""required"">「電話番号」を入力してください。<br /></span>" & vbCrLf
		If Len(Request.Form("phone1"))= 0 Then
			ErrMsg6 = ErrMsg6 & "<span class=""required"">「電話番号（市外局番）」を入力してください。<br /></span>" & vbCrLf
			ErrMsg6_1 = ErrMsg6_1 & "<span class=""required"">「電話番号（市外局番）」を入力してください。<br /></span>" & vbCrLf
		End If

		If Len(Request.Form("phone2"))= 0 Then
			ErrMsg6 = ErrMsg6 & "<span class=""required"">「電話番号（市内局番）」を入力してください。<br /></span>" & vbCrLf
			ErrMsg6_2 = ErrMsg6_1 & "<span class=""required"">「電話番号（市内局番）」を入力してください。<br /></span>" & vbCrLf
		End If

		If Len(Request.Form("phone3"))= 0 Then
			ErrMsg6 = ErrMsg6 & "<span class=""required"">「電話番号（加入者局番）」を入力してください。<br /></span>" & vbCrLf
			ErrMsg6_3 = ErrMsg6_3 & "<span class=""required"">「電話番号（加入者局番）」を入力してください。<br /></span>" & vbCrLf
		End If

	Else
		'市外局番（1か所目）
		intErrCheck = 0
'		strWords = "1234567890-"
		strWords = "1234567890"
		For i=1 To Len(Request.Form("phone1"))
			For j=1 To Len(strWords)
				If Mid(Request.Form("phone1"),i,1)=Mid(strWords,j,1) Then
					intErrCheck = intErrCheck + 1
				End If
			Next
		Next
		If intErrCheck <> Len(Request.Form("phone1")) Then
			ErrMsg6= ErrMsg6 & "<span class=""required"">「電話番号（市外局番）」は半角数字で入力してください。<br /></span>"
			ErrMsg6_1= ErrMsg6_1 & "<span class=""required"">「電話番号（市外局番）」は半角数字で入力してください。<br /></span>"
		End If

		'市内局番（2か所目）
		intErrCheck = 0
		For i=1 To Len(Request.Form("phone2"))
			For j=1 To Len(strWords)
				If Mid(Request.Form("phone2"),i,1)=Mid(strWords,j,1) Then
					intErrCheck = intErrCheck + 1
				End If
			Next
		Next
		If intErrCheck <> Len(Request.Form("phone2")) Then
			ErrMsg6= ErrMsg6 & "<span class=""required"">「電話番号（市内局番）」は半角数字で入力してください。<br /></span>"
			ErrMsg6_2= ErrMsg6_2 & "<span class=""required"">「電話番号（市内局番）」は半角数字で入力してください。<br /></span>"
		End If

		'加入者局番（3か所目）
		intErrCheck = 0
		For i=1 To Len(Request.Form("phone3"))
			For j=1 To Len(strWords)
				If Mid(Request.Form("phone3"),i,1)=Mid(strWords,j,1) Then
					intErrCheck = intErrCheck + 1
				End If
			Next
		Next
		If intErrCheck <> Len(Request.Form("phone3")) Then
			ErrMsg6= ErrMsg6 & "<span class=""required"">「電話番号（加入者局番）」は半角数字で入力してください。<br /></span>"
			ErrMsg6_3= ErrMsg6_3 & "<span class=""required"">「電話番号（加入者局番）」は半角数字で入力してください。<br /></span>"
		End If
	End If

'---　Email　---
	If Len(Request.Form("email"))= 0 Or Len(Request.Form("email2"))= 0 Then
		ErrMsg7 = ErrMsg7 & "<span class=""required"">「Email」を入力してください。<br /></span>" & vbCrLf
		If Len(Request.Form("email"))= 0 Then
			ErrMsg7_1 = ErrMsg7_1 & "<span class=""required"">「Email」を入力してください。<br /></span>" & vbCrLf
		End If
		
		If Len(Request.Form("email2"))= 0 Then
			ErrMsg7 = ErrMsg7 & "<span class=""required"">「確認用Email」を入力してください。<br /></span>" & vbCrLf
			ErrMsg7_2 = ErrMsg7_2 & "<span class=""required"">「Email」を入力してください。<br /></span>" & vbCrLf
		End If

	Else
		'Email
		intErrCheck = 0
		strWords = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890.@_-"
		For i=1 To Len(Request.Form("email"))
			For j=1 To Len(strWords)
				If Mid(Request.Form("email"),i,1)=Mid(strWords,j,1) Then
					intErrCheck = intErrCheck + 1
				End If
			Next
		Next


			mch = objBP.Match("m/^[\w\.\-]+@[\w_\-]+\.[\w_\.\-]*[a-z][a-z]+$/k",Request.Form("email"))
			If intErrCheck <> Len(Request.Form("email")) Then
				ErrMsg7 = ErrMsg7 & "<span class=""required"">「Email」に不正な文字が使われています。半角英数文字（.@_-の記号を含む）を入力してください。<br /></span>" & vbCrLf
				ErrMsg7_1 = ErrMsg7_1 & "<span class=""required"">「Email」に不正な文字が使われています。半角英数文字（.@_-の記号を含む）を入力してください。<br /></span>" & vbCrLf
				If mch = 0 Then
					ErrMsg7 = ErrMsg7 & "<span class=""required"">「Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>" & vbCrLf
					ErrMsg7_1 = ErrMsg7_1 & "<span class=""required"">「Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>" & vbCrLf
				End If
			ElseIf mch = 0 Then
				ErrMsg7 = ErrMsg7 & "<span class=""required"">「Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>"
				ErrMsg7_1 = ErrMsg7_1 & "<span class=""required"">「Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>"
			End If

		'確認用
		intErrCheck = 0
		For i=1 To Len(Request.Form("email2"))
			For j=1 To Len(strWords)
				If Mid(Request.Form("email2"),i,1)=Mid(strWords,j,1) Then
					intErrCheck = intErrCheck + 1
				End If
			Next
		Next

			mch = objBP.Match("m/^[\w\.\-]+@[\w_\-]+\.[\w_\.\-]*[a-z][a-z]+$/k",Request.Form("email2"))
			If intErrCheck <> Len(Request.Form("email2")) Then
				ErrMsg7 = ErrMsg7 & "<span class=""required"">「確認用Email」に不正な文字が使われています。半角英数文字（.@_-の記号を含む）を入力してください。<br /></span>" & vbCrLf
				ErrMsg7_2 = ErrMsg7_2 & "<span class=""required"">「確認用Email」に不正な文字が使われています。半角英数文字（.@_-の記号を含む）を入力してください。<br /></span>" & vbCrLf
				If mch = 0 Then
					ErrMsg7 = ErrMsg7 & "<span class=""required"">「確認用Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>" & vbCrLf
					ErrMsg7_2 = ErrMsg7_2 & "<span class=""required"">「確認用Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>" & vbCrLf
				End If
			ElseIf mch = 0 Then
				ErrMsg7 = ErrMsg7 & "<span class=""required"">「確認用Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>"
				ErrMsg7_2 = ErrMsg7_2 & "<span class=""required"">「確認用Email」の形式が不適切です。(例)kenko001@oki.com のように入力してください。<br /></span>"
			End If

		'Emailと確認用の一致確認
		comp = StrComp(Request.Form("email"),Request.Form("email2"))
		If comp <> 0 Then
			ErrMsg7 = ErrMsg7 & "<span class=""required"">「Email」と「確認用Email」は同じ内容を入力してください。<br /></span>" & vbCrLf
		End If
	End If

'---　メール内容　---
	If Len(Request.Form("note"))= 0 Then
		ErrMsg8= "<span class=""required"">「メール内容」を入力してください。<br /></span>"
	ElseIf Len(Request.Form("note"))> 175 Then
		ErrMsg8= ErrMsg8 & "<span class=""required"">「メール内容」は175文字以内でご入力ください。<br /></span>"
	End If

	'---　署名　---
	If Len(Request.Form("note"))= 0 Then
		ErrMsg8= "<span class=""required"">「署名」を入力してください。<br /></span>"
	ElseIf Len(Request.Form("note"))> 525 Then
		ErrMsg8= ErrMsg8 & "<span class=""required"">「メール内容」は175文字以内でご入力ください。<br /></span>"
	End If





ErrFlag= 0
	If Len(ErrMsg1 & ErrMsg2 & ErrMsg3 & ErrMsg4 & ErrMsg5 & ErrMsg6 & ErrMsg7 & ErrMsg8 & ErrMsg9)> 0 Then 'エラーがある場合
		ErrFlag= 1
	End If
End If

If Request.Form("SendFlag")= 1 And ErrFlag= 0 Then
	Server.Transfer("confirm.asp")
End If

%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>年賀状送信プログラム<% If ErrFlag= 1 Then Response.Write "（ERROR）" End If %>｜</title>
<meta name="viewport" content="width=device-width, initial-scale=1">


<link rel="stylesheet" type="text/css" href="/common/css/base.css">
<link rel="stylesheet" type="text/css" href="/common/css/contents.css">
<link rel="stylesheet" type="text/css" href="/common/css/responsive.css">
<link rel="stylesheet" type="text/css" href="https://use.fontawesome.com/releases/v6.0.0/css/all.css">
<link rel="stylesheet" type="text/css" href="/resource/lightbox.css">
<style type="text/css">
<!--
/* ----------------------------------------------------------------
    formStep
----------------------------------------------------------------- */
#formStep {
border: 1px solid #CCC;
font-size: 80%;
color: #333;
margin-bottom: 20px;
padding-top: 10px;
zoom:1;
background-repeat: no-repeat;
background-position: 294px 36px;
padding-bottom: 9px;
overflow:hidden;
}

#formStep ul {
position:relative;
left:50%;
float:left;
margin-left:-5px;
}

#formStep ul li {
position:relative;
left:-50%;
float: left;
margin-left:10px;
width: 227px;
}

#formStep .stepOn {
background-color: #FFECEC;
border-top-width: 1px;
border-right-width: 1px;
border-bottom-width: 2px;
border-left-width: 1px;
border-top-style: solid;
border-right-style: solid;
border-bottom-style: solid;
border-left-style: solid;
border-top-color: #CCC;
border-right-color: #CCC;
border-bottom-color: #DC0000;
border-left-color: #CCC;
margin-left: 40px;
padding-left: 10px;
}
#formStep .stepOff {
background-color: #F7F7F7;
margin-left: 40px;
padding-left: 10px;
border: 1px solid #CCC;
}
#formStep .stepNumber {
font-size: 150%;
overflow: visible;
line-height: 120%;
font-weight: bold;
padding-right: 5px;
}

/* ----------------------------------------------------------------
    table
----------------------------------------------------------------- */
th {
white-space:nowrap;
}

.itemComment {
margin-bottom:0 !important;
padding-bottom:0 !important;
}

/* ----------------------------------------------------------------
    input/type=text,　textarea
----------------------------------------------------------------- */
input[type="text"][size="70"] {
width:50%;
}

input[type="text"][size="10"] {
width:20%;
}

input[type="text"][size="3"] {
width:20%;
}

input[type="text"][size="4"] {
width:20%;
}

/* ----------------------------------------------------------------
    submit/reset
----------------------------------------------------------------- */
#inquiryForm {
text-align: center;
}
#inquiryForm .submitArea input[type="reset"] + input[type="submit"] {
margin-left: 30px;
}
#inquiryForm .submitArea {
text-align: center;
margin-top: 20px;
margin-bottom: 20px;
}
#inquiryForm .submitArea form {
display: inline;
}
#inquiryForm .submitArea form + form {
margin-left: 30px;
}

/* ----------------------------------------------------------------
    clearfix
----------------------------------------------------------------- */
.clearfix {
clear: both;
}

/* ----------------------------------------------------------------
   1024px以下
----------------------------------------------------------------- */
@media screen and (max-width:1024px) {
#formStep ul li {
width:150px;
}

#formStep .stepOn {
margin-left: 10px;
padding-left: 10px;
}

#formStep .stepOff {
margin-left: 10px;
padding-left: 10px;
}
}

/* ----------------------------------------------------------------
   600px以下
----------------------------------------------------------------- */
@media screen and (max-width:600px) {
#formStep ul li {
width:120px;
}

#formStep .stepOn {
margin-left: 10px;
padding-left: 10px;
}

#formStep .stepOff {
margin-left: 10px;
padding-left: 10px;
}

td {
width:200px;
}

.itemComment {
width:200px;
}

textarea {
width:200px;
}
}
-->
</style>
<script type="text/javascript" src="/resource/lightbox_plus_min.js"></script>
<script type="text/javascript" src="https://code.jquery.com/jquery-3.4.1.min.js"></script>
<script>
$(function(){
var _window = $(window),
_header = $('.site-header'),
heroBottom;

_window.on('scroll',function(){
heroBottom = $('.hero').height();
if(_window.scrollTop() > heroBottom){
_header.addClass('transform');   
}
else{
_header.removeClass('transform');   
}
});

_window.trigger('scroll');	
});
</script>
<script>
$(function() {
var topBtn = $('#page_top');	
topBtn.hide();
$(window).scroll(function () {
if ($(this).scrollTop() > 100) {
topBtn.fadeIn();
} else {
topBtn.fadeOut();
}
});
//スクロールしてトップ
topBtn.click(function () {
$('body,html').animate({
scrollTop: 0
}, 500);
return false;
});
});

$(function() {
  $('#menu li').hover(function() {
    $(this).find('.menu_contents').stop().slideDown();
  }, function() {
    $(this).find('.menu_contents').stop().slideUp();
 
  });
 
});
</script>
<script type="text/javascript" src="/common/js/hovermenu.js"></script>
<script type="text/javascript" src="cmn/js/template.js"></script>
<script type="text/javascript" src="cmn/js/form.js"></script>
<script src="https://ajaxzip3.github.io/ajaxzip3.js" charset="UTF-8"></script>
</head>

<body class="under">
 
    <header>
</header>
    
    
   
   
    <div id="contents">


<h3>メール送信</h3>


<% If ErrFlag= 1 Then %>
<div id="errorMessage"><p>ご入力内容に誤りがあります。<img src="cmn/img/form/error_icon.gif" width="43" height="39" alt="Error icon" />マークの項目をご確認いただき、再度入力をお願いいたします。</p></div>
<% End If %>

<p><em class="attention01">年賀状送信のために以下の設定をしてください。</em></p>

<table>
	<!--　メール送信先リスト読込み　-->
<tr<% If Len(ErrMsg1)> 0 Then Response.Write " class=""inputError""" End If %>>
<th>
<label for="id1">メール送信先リスト読込み <span class="required">（必須）</span></label>
<% If Len(ErrMsg1)>0 Then %>
<% End If %>
  <button>メール送信先リスト読込み</button>


</th>

<!--　画像・動画読み込み　-->
<tr<% If Len(ErrMsg1)> 0 Then Response.Write " class=""inputError""" End If %>>
<th>
<label for="id1">画像・動画読み込み <span class="required">（必須）</span></label>
<% If Len(ErrMsg1)>0 Then %>
<% End If %>
<div class="filearea" id="video" ondrop="dropHandler(event);" ondragover="dragOverHandler(event);">
    <label>
      <input type="file"  onchange="preview(this, 'preview-video');">
      <p class="selecter"></p>
      <video class="preview pre-select" id="preview-video" controls="controls" width = 320 height = 180>
      <script>
       function preview(obj, previewId) {
    previewFile(obj.files, previewId);
}

/**
 * @param {File} files 入力ファイル
 * @param {String} previewId プレビュー表示DOMのID
 */
function previewFile(files, previewId) {
    let fileReader = new FileReader();
    fileReader.onload = (function () {
        document.getElementById(previewId).src = fileReader.result;
        document.getElementById(`${previewId}-file`).innerText = files[0].name;
    });
    fileReader.readAsDataURL(files[0]);
}

/** 
 * @param {Event} event
 */
function dropHandler(event) {

    event.preventDefault();

    if (event.dataTransfer.files.length === 0) return false;
    const files = event.dataTransfer.files;
    const previewId = `preview-${event.currentTarget.id}`;
    previewFile(files, previewId);
}
      </script>
      </video>
    </label>
  </div>


<form method="post" action="index.asp" id="inquiryForm" name="inquiryForm">
<table class="tb03" id="outline1">

<!--　件名　-->
<tr<% If Len(ErrMsg1)> 0 Then Response.Write " class=""inputError""" End If %>>
<th>
<label for="id1">件名 <span class="required">（必須）</span></label>
<% If Len(ErrMsg1)>0 Then %>
<div class="errorIcon"><img src="cmn/img/form/error_icon.gif" width="43" height="39" alt="Error icon" /></div>
<% End If %>
</th>
<td>
<%=ErrMsg1 %>
<input type="text" name="name" id="id1" size="70" maxlength="35" value="<%=Request.Form("name") %>"<% If Len(ErrMsg1)> 0 Then Response.Write " class=""formOnError""" End If %> />
</td>
</tr>


<!--　送信元アドレス　-->
<tr>
<th<% If Len(ErrMsg3)> 0 Then Response.Write " class=""inputError""" End If %>>
<label for="id3">送信元アドレス <span class="required">（必須）</span></label>
<% If Len(ErrMsg3)>0 Then %>
<div class="errorIcon"><img src="cmn/img/form/error_icon.gif" width="43" height="39" alt="Error icon" /></div>
<% End If %>
</th>
<td>
<%=ErrMsg3 %>
<input type="text" name="company" id="id3" size="70" maxlength="254" value="<%=objBP.HAN2ZEN(CStr(Request.Form("company"))) %>"<% If Len(ErrMsg3)> 0 Then Response.Write " class=""formOnError""" End If %> />
</td>
</tr>

<!--　メール内容　-->
<tr<% If Len(ErrMsg8)> 0 Then Response.Write " class=""inputError""" End If %>>
<th>
<label for="id8">メール内容 <span class="required">（必須）</span></label>
<% If Len(ErrMsg8)>0 Then %>
<div class="errorIcon"><img src="cmn/img/form/error_icon.gif" width="43" height="39" alt="Error icon" /></div>
<% End If %>
</th>
<td>
<p class="itemComment">カタカナは全角でご入力ください。（175文字以内）</p>
<% If Len(ErrMsg8)>0 Then %>
<br />
<%=ErrMsg8 %>
<% End If %>
<textarea name="note" id="id8" cols="70" rows="10"<% If Len(ErrMsg8)> 0 Then Response.Write " class=""formOnError""" End If %> ><%=objBP.HAN2ZEN(CStr(Request.Form("note"))) %></textarea>
</td>
</tr>



<!--　署名　-->
<tr<% If Len(ErrMsg8)> 0 Then Response.Write " class=""inputError""" End If %>>
<th>
<label for="id8">署名 <span class="required">（必須）</span></label>
<% If Len(ErrMsg8)>0 Then %>
<div class="errorIcon"><img src="cmn/img/form/error_icon.gif" width="43" height="39" alt="Error icon" /></div>
<% End If %>
</th>
<td>
<% If Len(ErrMsg8)>0 Then %>
<br />
<%=ErrMsg8 %>
<% End If %>
<textarea name="note" id="id8" cols="70" rows="10"<% If Len(ErrMsg8)> 0 Then Response.Write " class=""formOnError""" End If %> ><%=objBP.HAN2ZEN(CStr(Request.Form("note"))) %></textarea>
</td>
</tr>

</table>

<input type="hidden" name="SendFlag" value="1" />
<div class="submitArea">
<input type="button" onclick="location.href='kakunin.asp'" value="一斉送信">
</div>
</form>

</div>
<!--/#undercontents-->

</div>
<!--/#contents-->




<div id="page_top"><a href="#"></a></div>

<script src="/common/js/5-1-3.js"></script>
<script type="text/javascript">
(function($) {
  var $nav   = $('#navArea');
  var $btn   = $('.toggle_btn_under');
  var $mask  = $('#mask');
  var open   = 'open'; // class
  // menu open close
  $btn.on( 'click', function() {
    if ( ! $nav.hasClass( open ) ) {
      $nav.addClass( open );
    } else {
      $nav.removeClass( open );
    }
  });
  // mask close
  $mask.on('click', function() {
    $nav.removeClass( open );
  });
} )(jQuery);
</script>

</body>
</html>
<%
Set objBP = Nothing
If Err.Number <> 0 Then
'Response.Redirect "error.html"
End If
%>