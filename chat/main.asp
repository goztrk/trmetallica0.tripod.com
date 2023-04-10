<% Response.Redirect "main2.asp" %>

<html>

<head>
<title>ASP Chat</title>
<script language="JavaScript">

     function closeWin(){
           self.close();
     }

</script>
</head>
<frameset Marginwidth="0" Marginheight="0" Framespacing="0" Border="0" Frameborder="0"
rows="*,94">
  <frame Name="main-body" Src="main2.asp" Scrolling="auto" Marginwidth="0" Marginheight="0"
  Framespacing="0" Border="0" Frameborder="NO">
  <frame Name="bottom" Src="bottom.asp"  Marginwidth="0"
  Marginheight="0" Framespacing="0" Border="0" Frameborder="NO" scrolling=no>
</frameset>

</html>
