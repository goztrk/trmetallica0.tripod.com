<%

   User = Request.QueryString("user")
   From = Request.QueryString("from")

%>

<html>
<head>
<title>Gizli mesaj gönderme:<%=User%></title>

<script language="JavaScript">

  <% If User = "" Then %>
    window.close();
  <% End If %>

   function check(form){
          form.submit();
   }

</script>
</head>

<body BGCOLOR="#476895" TEXT=#f0f0f0 LINK=#f0f0f0 VLINK=#f0f0f0>

<center>
<font face=verdana size=2>
Gizli mesaj gönderilecek kiþi:<%=User%>
</font>
<form action=sendprivate.asp method=post>
<table border=0 cellspacing=0>
<tr><td valign=top><font face='verdana, arial' size=-1>Mesajýnýz:
</td><td valign=top>
  <textarea name="Msg" rows=3 cols=30></textarea>
</td></tr> 
<tr><td colspan=2 align=center>
<br><br>
<input type="button" value="Gönder" onClick="check(this.form)">
<input type="hidden" value="<%=User%>" name="User">
<input type="hidden" value="<%=From%>" name="From">
</td></tr></form></table>
</body>
</html>