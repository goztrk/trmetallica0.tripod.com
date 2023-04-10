<html>
<head>
<SCRIPT LANGUAGE="JavaScript">
<!--

   setTimeout("update()",15000);

   function update(){
         document.location.href="users.asp"
   }

   function privateWin(user, from){
         url = "privatemsg.asp?user=" + user + "&from=" + from;
         privatemsg = window.open(url,"privatemsg","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,WIDTH=400,HEIGHT=200");
   }
//-->

</SCRIPT>
</head>
<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 BGCOLOR="#ffffff" TEXT=#000000 LINK=#000000 VLINK=#000000>
<table width=186 border=0 cellspacing=0><tr><td bgcolor=#cccccc width=100% align=center>
<font face="Trebuchet MS, ms sans serif, Verdana,Arial" size=2 color=#476895><b>Odadakiler</b></font>
</td></tr></table>
<br>
<font face='verdana,arial' size=-1 0>
<%

   Topic = Session("Topic")
   
   Application.Lock
   

   If Not IsEmpty(Application.Contents("Users")) Then
   		 Users = Application("Users")
		 Group = Application("UsersGroup") 
         NumOfUsers = UBound(Users)
         For I = 0 To NumOfUsers
               If Topic = Group(I) Then
		   			Response.Write ""
                   	Response.Write "&nbsp;<a href=""javascript:privateWin('" & Server.URLEncode(Users(I)) & "', '" & Server.URLEncode(Session("Name")) & "')""><img align=absmiddle src='images/lips.gif' border=0> <b>" & Users(I) & "</b></a><br>"
		   			Response.Write "&nbsp;<small><b>IP:</b> " & Application("UsersIP")(I) & "<br>"
		   			Response.Write "&nbsp;<b>Giriþ:</b> " & Application("UsersTimeOn")(I) & "</small><br><br>"
               End If
         Next
   End If
   
   Application.Unlock


%>

<br><br><br>
<font face="ms sans serif" size=1 color=#476895><b>Gizli mesaj göndermek için ismi týklayýnýz. </b></font>

</body>
</html>