<%  Application("filebase") = Server.MapPath("files") & "\" %>


<html>
<head>
  <title>Sohbet Giriþi</title>
  <script language="JavaScript">
     function verify(form){
           if(form.Name.value == ""){
                alert("Lütfen isminizi giriniz!");
           }else{
           }
     }
  </script>
</head>
<body Bgcolor="#476895" TEXT=#f0f0f0>
<%=Request.QueryString("ErrorMessage")%>
<BR>
<center>
<CENTER>
<img src="images/chat.gif">
<br>
<font face="Trebuchet MS, ms sans serif, Verdana,Arial" size=2 color=#f0f0f0><b>Lütfen isminizi giriniz ve girmek istediðiniz odayý seçiniz.</b>
</center>
<BR>
<BR>    <center>
    <center>
      <form name="form_topics" action="chatroom.asp" method="post" onSubmit="verify(this)">
      <table>

           <tr>
              <td>
                  <font face="Trebuchet MS, ms sans serif, Verdana,Arial" size=2 color=#f0f0f0>Ýsim:
              </td>
              <td>
                  <%
                      If Not IsNull(Session("Name")) Then
                           Response.Write "<input type='text'  name='Name' value='" & Session("Name") & "' maxlength=12 maxsize=12>"
                      Else
                           Response.Write "<input type='text' name='Name' maxlength=12 maxsize=12>"
                      End If
                   %>
             </td>
           </tr><tr>
              <td>
                  <font size=2 face=verdana color=#f0f0f0>Oda:
              </td>
              <td>
                  <select name="Topic">

                          <%
                              Topics = Application("TopicNames")
                              If IsArray(Topics) Then
                                  For I=1 To UBound(Topics)
                                     Name = Topics(I)
                                     Response.Write "<option value='" & Name & "'>" & Name
                                  Next
                              End If
                          %>


                  </select>
             </td>
           </tr><tr>
             <td colspan=2 align=right><br>
                  <input type="image" src="images/logon.gif" border=0>
             </td>
           </tr>
      </table>
      </form>
    </center>
</body>
</html>