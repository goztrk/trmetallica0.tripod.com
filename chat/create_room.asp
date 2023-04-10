<%

   Topics = Application("TopicNames")
   If IsArray(Topics) Then
         NumOfTopics = UBound(Topics)
         ReDim Preserve Topics(NumOfTopics+1)
         Topics(NumOfTopics+1) = Request.Form("room_name")

         Descriptions = Application("TopicDescriptions")
         ReDim Preserve Descriptions(NumOfTopics+1)
         Descriptions(NumOfTopics+1) = Request.Form("room_description")
         Application("TopicNames") = Topics
         Application("TopicDescriptions") = Descriptions
         Response.Redirect("default.asp")
    Else
         ReDim Topics(1)
         Topics(0) = Request.Form("room_name")

         ReDim Descriptions(1)
         Descriptions(1) = Request.Form("room_description")

         Application("TopicNames") = Topics
         Application("TopicDescriptions") = Descriptions
         Response.Redirect("default.asp")
    End If


%>