<HTML>
<HEAD>

<script language="javascript">

   function handle(form){
       var text = form.visMsg.value
       if(text == "") return false;
       form.Msg.value = text;
       form.RefreshRate.value = document.forms[1].intreval[document.forms[1].intreval.selectedIndex].value
       form.visMsg.value="";
       form.visMsg.focus();
       return true;
   }

  function exitChat(){
       parent.frames.location.href='exit.asp';
       return false;
  }

</script>

</HEAD>
<BODY BGCOLOR="cccccc" TEXT=#000000 onLoad="document.message.visMsg.focus()">

<TABLE><TR><TD  ALIGN=CENTER VALIGN=TOP>
<FORM NAME="message" TARGET="chat" ACTION="chat.asp" METHOD="POST" onSubmit="return handle(this)">
<font face="ms sans serif" size=2>Mesaj:</font>
<input type="text" size=40 name="visMsg">
<input type="hidden" Name="Msg" value="">
<input type="hidden" Name="RefreshRate">
<input type="hidden" Name="Add" value="TRUE">
</TD>
</form>
<form onSubmit="return exitChat()">
<td valign=top>
<font face='ms sans serif' size=2 color=#000000>
Yenileme aralýðý: <select name="intreval">
<option value="10000">10 sn.
<option value="20000">20 sn.
<option value="30000">30 sn.
<option value="45000">45 sn.
</select>
</td><td valign=bottom>
<INPUT type="image" border=0 src="images/logoff.gif">


</TD></TR><tr><td> <font face="Trebuchet MS, ms sans serif, Verdana,Arial" size=1 color=#476895><b> Copyright by Adrenalin Labs. / Çeviren <a href="http://www.okyanus.gen.tr" target="okyanus">Syntact``</a></b></font></td><td></td><td></td></tr></TABLE>
</CENTER></FORM>
</BODY>
</HTML>