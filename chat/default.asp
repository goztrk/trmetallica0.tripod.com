
<html>
<head>
  <Title>ASP Chat (T�rk�e versiyon)</title>

<SCRIPT LANGUAGE="JavaScript">

  function enter(){
        chatWin = window.open("main.asp","chatWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,WIDTH=630,HEIGHT=540");
        chatWin.location.href = "main.asp"
        if(navigator.appVersion.indexOf("4") != -1){
            chatWin.moveTo(0,0);
        }
  }

     function closeWin(){
           chatWin.close();
     }
</SCRIPT>
</head>

<body  Bgcolor="#476895" link=#000000 vlink=#000000>
<CENTER>
<img src="images/chat.gif">
</center><br>
<center>
<table border="0" cellpadding="0" cellspacing="0">
<tr><td align="center">
Bir sohbet odas�na girmek i�in 'Giri�' d��mesini t�klay�n�z. E�er kendi sohbet odan�z� yaratmak istiyorsan�z <a href="create_room.html">buray� t�klay�n�z</a>.
<br><br>
</td></tr>
<tr><td align=right>
<form onSubmit="enter()"><input type=image src="images/enter.gif" border=0 value="ASP Chat`e giri�" ></form>
</td></tr>
</table>







