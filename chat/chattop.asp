

<html>
<head>
<script language="JavaScript">

//window.onunload=exit

function exit(){
     //alert("Called");
     top.location.href="exit.asp";
}
</script>
</head>
  <body bgcolor=#cccccc text=#476895>
     <font face="Trebuchet MS, ms sans serif, Verdana,Arial" size=3 color=#476895><b>Sohbet odasý: <%=Request("Topic")%></b>
	&nbsp;&nbsp;&nbsp; <font size=2>Kullandýðýnýz isim: <%=Session("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</font>

  </body>
</html>