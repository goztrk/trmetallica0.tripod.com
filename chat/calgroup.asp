<% 

    GroupName = Request("Topic")
    UserName = Request("Name")

   Topics = Application("TopicNames")
   If IsArray(Topics) Then
         NumOfTopics = UBound(Topics)
         For I = 0 To NumOfTopics
             If Topics(I) = GroupName Then
		'Room already exists, just redirect to it.
                Response.Redirect "chatroom.asp?Name=" & Server.URLEncode(UserName) & "&Topic=" & Server.URLEncode(GroupName)
             End If
         Next

         ReDim Preserve Topics(NumOfTopics+1)
         Topics(NumOfTopics+1) = GroupName

         Descriptions = Application("TopicDescriptions")
         ReDim Preserve Descriptions(NumOfTopics+1)
         Descriptions(NumOfTopics+1) = ""
         Application("TopicNames") = Topics
         Application("TopicDescriptions") = Descriptions
                Response.Redirect "chatroom.asp?Name=" & Server.URLEncode(UserName) & "&Topic=" & Server.URLEncode(GroupName)
    Else
         ReDim Topics(1)
         Topics(0) = GroupName

         ReDim Descriptions(1)
         Descriptions(1) = GroupDescription

         Application("TopicNames") = Topics
         Application("TopicDescriptions") = Descriptions
         Response.Redirect "chatroom.asp?Name=" & Server.URLEncode(UserName) & "&Topic=" & Server.URLEncode(GroupName)
    End If
%>