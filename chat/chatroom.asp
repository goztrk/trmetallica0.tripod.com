<%
   If Request("Name") = "" OR Request("Topic") = "" Then
        Response.Redirect("default.asp?0")
   Else
        Topic = FilterTopic(Request("Topic"))
        Session("Topic") = Topic
        Session("Name") = Request("Name")
   End If

   Name = Request("Name")
   Topic = Request("Topic")

   Const BadName = "<font face=verdana size=+1><b>Seçtiðiniz isim kullanýmda , baþka bir isim seçiniz.</b></font><br><br>"

   If Not IsEmpty(Application("Users")) Then
      If IsArray(Application("Users")) Then
          uLen = UBound(Application("Users"))
          For I = 0 To uLen
               If Application("Users")(I) = Name Then
		   			Name = Name & "_1"
        	   		Session("Name") = Name
                   'Response.Redirect("main2.asp?ErrorMessage=" & Server.URLEncode(BadName))
		   			Exit For
               End If
          Next
          Users = Application("Users")
          UsersGroup = Application("UsersGroup")
          NumOfUsers = UBound(Users)+1
          ReDim Preserve Users(NumOfUsers)
          ReDim Preserve UsersGroup(NumOfUsers)
          'Response.AppendToLog "Number of users when " & Name & " entered: " & NumOfUsers
          Users(NumOfUsers-1) = Name
          UsersGroup(NumOfUsers-1) = Topic
          Application("Users") = Users
          Application("UsersGroup") = UsersGroup


      	  UsersIP = Application("UsersIP")
          UsersIdleTime = Application("UsersIdleTime")
          UsersTimeOn = Application("UsersTimeOn")
          ReDim Preserve UsersIP(NumOfUsers)
          ReDim Preserve UsersIdleTime(NumOfUsers)
          ReDim Preserve UsersTimeOn(NumOfUsers)
      	  UsersTimeOn(NumOfUsers-1) = Now
          UsersIdleTime(NumOfUsers-1) = 0
          UsersIP(NumOfUsers-1) = Request.ServerVariables("REMOTE_ADDR")
          Application("UsersIP") = UsersIP
          Application("UsersTimeOn") = UsersTimeOn
          Application("UsersIdleTime") = UsersIdleTime
      Else
          'Response.AppendToLog "Number of users when " & Name & " entered: " & "0 (Not Array)"
          ReDim Users(1)
          ReDim UsersGroup(1)
          Users(0) = Name
          UsersGroup(0) = Topic
          Application("Users") = Users
          Application("UsersGroup") = UsersGroup

          Redim UsersIP(1)
          Redim UsersIdleTime(1)
          Redim UsersTimeOn(1)
          UsersTimeOn(0) = Now
          UsersIdleTime(0) = 0
          UsersIP(0) = Request.ServerVariables("REMOTE_ADDR")
          Application("UsersIP") = UsersIP
          Application("UsersTimeOn") = UsersTimeOn
          Application("UsersIdleTime") = UsersIdleTime

      End If
   Else
     'Initialize Users Variable
      Dim NewUsers()
      Redim NewUsers(1)
      Dim NewUsersGroup()
      Redim NewUsersGroup(1)
      NewUsers(0) = Name
      NewUsersGroup(0) = Topic
      Application("Users") = NewUsers
      Application("UsersGroup") = NewUsersGroup


      Dim NewUsersIP()
      Redim NewUsersIP(1)
      Dim NewUsersIdleTime()
      Redim NewUsersIdleTime(1)
      Dim NewUsersTimeOn()
      Redim NewUsersTimeOn(1)
      NewUsersTimeOn(0) = Now
      NewUsersIdleTime(0) = 0
      NewUsersIP(0) = Request.ServerVariables("REMOTE_ADDR")
      Application("UsersIP") = NewUsersIP
      Application("UsersTimeOn") = NewUsersTimeOn
      Application("UsersIdleTime") = NewUsersIdleTime
   End If
   Session("Enter") = True

   Function FilterTopic(Topic)
        Topic = Replace(Topic, ":", ".")
        Topic = Replace(Topic, "/", ".")
        Topic = Replace(Topic, "\", ".")
	FilterTopic = Topic
   End Function

%>

<HTML>
<HEAD>
<TITLE><%=Topic%> Sohbet Odasý</TITLE>


</HEAD>
      <frameset  Framespacing="0" Border="0" Frameborder="0" rows="30,*, 50">
        <frame name="chattop" scrolling="no" marginwidth=8 marginheight=0 noresize src="chattop.asp?Topic=<%=Server.URLEncode(Topic)%>" >
          <frameset marginwidth="0" merginheight="0" framespacing="0" Border="1" frameborder="1" cols="*,186">
              <frame name="chat" scrolling="yes" marginwidth=8 marginheight=0 noresize src="chat.asp">
              <frame name="users" scrolling="auto" marginwidth=2 marginheight=0  src="users.asp">
          </frameset>
        <frame Name="bottom" Src="msg.asp"  Marginwidth="0"
           Marginheight="0" Framespacing="0" Border="0" Frameborder="NO" scrolling=no>
      </frameset>

</frameset>


</HTML>