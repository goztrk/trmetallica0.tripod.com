<%

    Topic = Session("Topic")
    Name = Session("Name")

    Application.Lock
	
    Users = Application("Users")
    UsersIP = Application("UsersIP")
    UsersTimeOn = Application("UsersTimeOn")
    UsersIdleTime = Application("UsersIdleTime")

    Match = 1
    If IsArray(Users) Then
		'   Response.Write Name & "<br>IsArray=True<br>"
           UsersGroup = Application("UsersGroup")
           NumOfUsers = UBound(Users)
           If NumOfUsers = 1 Then
		   		  '  Response.Write "UBound(Users)=1<br>"
                 	Application("Users") = ""
                 	Application("UsersGroup") = ""

                 	Application("UsersIP") = ""
                 	Application("UsersIdleTime") = ""
                 	Application("UsersTimeOn") = ""
           Else 
              ReDim UpdatedUsers(1)
              ReDim UpdatedUsersGroup(1)

              ReDim UpdatedUsersIP(1)
              ReDim UpdatedUsersTimeOn(1)
              ReDim UpdatedUsersIdleTime(1)
              For I = 0 To NumOfUsers-1
			  	 ' Response.Write "At " & I & " User is:" & Users(I) & "<br>"
                  If Name <> Users(I) Then
                       ReDim Preserve UpdatedUsers(Match)
                       ReDim Preserve UpdatedUsersGroup(Match)
                       UpdatedUsers(Match-1) = Users(I)
                       UpdatedUsersGroup(Match-1) = UsersGroup(I)

                       ReDim Preserve UpdatedUsersIP(Match)
                       ReDim Preserve UpdatedUsersTimeOn(Match)
                       ReDim Preserve UpdatedUsersIdleTime(Match)
                       UpdatedUsersIP(Match-1) = UsersIP(I)
                       UpdatedUsersTimeOn(Match-1) = UsersTimeOn(I)
                       UpdatedUsersIdleTime(Match-1) = UsersIdleTime(I)

                       Match = Match + 1
                  End If
              Next
              Application("Users") = UpdatedUsers
              Application("UsersGroup") = UpdatedUsersGroup

              Application("UsersIP") = UpdatedUsersIP
              Application("UsersTimeOn") = UpdatedUsersTimeOn
              Application("UsersIdleTime") = UpdatedUsersIdleTime

           End If
   End If
   
   Application.UnLock


   Set fileObject = Server.CreateObject("Scripting.FileSystemObject")
   textFile = Server.MapPath("files/" & Replace(Topic, "+", "_") & ".txt")
   Set inStream = fileObject.OpenTextFile(textFile,1,TRUE,FALSE)
   file = inStream.ReadAll
   Set inStream = Nothing

   file = file & "<FONT SIZE=2 FACE='Vedana,Arial' Color=#808080><b>***[" &  Name & " odadan ayrýldý (" & Now & ")]</b></FONT><br>"
   Set outStream = fileObject.CreateTextFile(textFile,True)
   outStream.WriteLine(file)
   outStream.Close
   Set outStream = Nothing


   '################################
   'Should just call Session.Abandon
   '################################
    'Session.Abandon


%>


<html>
<head>
  <script language="javascript">
      function closeWin(){
          if(opener){
               self.close();
          }else{
               //customize this to go where you want
               window.document.location.href="default.asp";
          }
      }
  </script>
</head>
  <body onLoad="closeWin()">
  </body>
</html>
