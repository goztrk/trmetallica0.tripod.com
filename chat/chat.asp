<%
   Response.Expires=-1000
   Topic = Session("Topic")
   Name = Session("Name") 
   If Request("RefreshRate") <> Empty Then
	Session("RefreshRate") = Request("RefreshRate")
   Else
	Session("RefreshRate") = Session("RefreshRate")
   End If

   If Session("RefreshRate") = "" Then
	Session("RefreshRate") = 10000
   End If
%>

<HTML>
<head>
<SCRIPT LANGUAGE="JavaScript">
<!--



   function update(){
         document.location.href="chat.asp"

   }

   function toBottom(){
	if(document.all){
		window.scroll(0, document.body.clientHeight);
	}else{
		window.scroll(0,20000)
	}
   }

   setTimeout("update()",<%=Session("RefreshRate")%>);
//-->

</SCRIPT>
</head>
<BODY BGCOLOR="#ffffff" onLoad="toBottom()">
<br>
<FONT SIZE=2 FACE="Verdana,Arial">
<%
     Set fileObject = Server.CreateObject("Scripting.FileSystemObject")
     'Map appropriate topic file
     textFile = Server.MapPath("files/" & Replace(Topic, "+", "_") & ".txt")
     If Not fileObject.FileExists(textfile) Then
          Set inStream = fileObject.OpenTextFile(textFile,8,True,FALSE)
          inStream.WriteLine " "
     End If

     Set inStream = fileObject.OpenTextFile(textFile,1,TRUE,FALSE)
	 

     file = inStream.ReadAll
     Set inStream = Nothing

     If Len(file) > 5000 Then
	 'archive the old file
	 'archive name is topic_date put under files/archives
	 archiveFile = Server.MapPath("files/" & Replace(Topic, "+", "_") & "_" & Server.URLEncode(Now) & ".txt")
         Set inStream = fileObject.CreateTextFile(archiveFile,True)
         inStream.Write file
	 inStream.Close
         Set inStream = Nothing


         Set inStream = fileObject.CreateTextFile(textFile,True)
         inStream.Write " "
         Set inStream = Nothing
         file = " "
     End If

     If IsNull(Session("Enter")) Or Session("Enter") = True Then
         file =  file & "<FONT SIZE=2 FACE='Vedana,Arial' Color=#808080><b>***[" &  name & " odaya girdi (" & Now & ")]***</b></FONT><br>"
         Session("Enter") = False
     End If

     If Request.Form("Add") = "TRUE" Then
        file =  file & "<FONT SIZE=2 Color=#000000 FACE='Verdana,Arial'><b>" & name & "</b><small>(" & Now & ")</small>:</FONT>" & Request("Msg") & "<br>" 
     End If

     Set outStream = fileObject.CreateTextFile(textFile,True)
     outStream.WriteLine(file)
     Set outStream = Nothing

     Response.Write "<font color=#000000>" &  file & ""

    PrivRcpt = Application("PrivateRcpt")
    Match = 0
    If IsArray(PrivRcpt) Then
           PrivateMsgs = Application("PrivateMsgs")
           NumOfRcpt = UBound(PrivRcpt)
           ReDim Preserve UpdatedPrivRcpt(1)
           ReDim Preserve UpdatedPrivMsg(1)
           For I = 0 To NumOfRcpt
                  If Name = PrivRcpt(I) Then
                       Response.Write "<FONT color=#ff0000 SIZE=2 FACE='Verdana,Arial'><b>*** " & PrivateMsgs(I) & " ***</b></font>"
                  Else
                       ReDim Preserve UpdatedPrivRcpt(Match)
                       ReDim Preserve UpdatedPrivMsg(Match)
                       UpdatedPrivRcpt(Match) = PrivRcpt(I)
                       UpdatedPrivMsg(Match) = PrivateMsgs(I)
                       Match = Match + 1
                  End If
              Next
              Application("PrivateRcpt") = UpdatedPrivRcpt
              Application("PrivateMsgs") = UpdatedPrivMsg
   End If

%>
</FONT>
<br>
<br>


<br>

</BODY>
</HTML>