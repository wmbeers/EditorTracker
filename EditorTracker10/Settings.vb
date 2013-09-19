Imports System.Xml

Namespace My

    'This class allows you to handle specific events on the settings class:
    ' The SettingChanging event is raised before a setting's value is changed.
    ' The PropertyChanged event is raised after a setting's value is changed.
    ' The SettingsLoaded event is raised after the setting values are loaded.
    ' The SettingsSaving event is raised before the setting values are saved.
    Partial Friend NotInheritable Class MySettings

        Public Sub SaveTo(ByVal path As String)
            Try
                Dim ConfigDoc As XmlDocument = New XmlDocument()
                Dim sNode, node As XmlElement
                node = ConfigDoc.AppendChild(ConfigDoc.CreateElement("configuration"))
                sNode = node.AppendChild(ConfigDoc.CreateElement("EditorTrackerSettings"))
                node = sNode.AppendChild(ConfigDoc.CreateElement("add"))
                node.SetAttribute("key", "EditorNameField")
                node.SetAttribute("value", EditorNameField)
                node = sNode.AppendChild(ConfigDoc.CreateElement("add"))
                node.SetAttribute("key", "EditDateField")
                node.SetAttribute("value", EditDateField)
                node = sNode.AppendChild(ConfigDoc.CreateElement("add"))
                node.SetAttribute("key", "TrackDatabaseUser")
                node.SetAttribute("value", TrackDatabaseUser)
                node = sNode.AppendChild(ConfigDoc.CreateElement("add"))
                node.SetAttribute("key", "DateAndTime")
                node.SetAttribute("value", DateAndTime)
                node = sNode.AppendChild(ConfigDoc.CreateElement("add"))
                node.SetAttribute("key", "X")
                node.SetAttribute("value", XField)
                node = sNode.AppendChild(ConfigDoc.CreateElement("add"))
                node.SetAttribute("key", "Y")
                node.SetAttribute("value", YField)
                node = sNode.AppendChild(ConfigDoc.CreateElement("add"))
                node.SetAttribute("key", "USNG")
                node.SetAttribute("value", USNGField)
                ConfigDoc.Save(path)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(ex.ToString)
            End Try
        End Sub

        Public Sub LoadFrom(ByVal path As String)
            Try
                Dim ConfigDoc As XmlDocument = New XmlDocument()
                ConfigDoc.Load(path)
                For Each node In ConfigDoc.SelectNodes("/configuration/EditorTrackerSettings/add")
                    Dim el As XmlElement = TryCast(node, XmlElement)
                    If el Is Nothing Then Continue For
                    If el.GetAttribute("key") = "EditorNameField" Then
                        EditorNameField = el.GetAttribute("value")
                    ElseIf el.GetAttribute("key") = "EditDateField" Then
                        EditDateField = el.GetAttribute("value")
                    ElseIf el.GetAttribute("key") = "TrackDatabaseUser" Then
                        Dim bln As Boolean
                        If Boolean.TryParse(el.GetAttribute("value"), bln) Then TrackDatabaseUser = bln
                    ElseIf el.GetAttribute("key") = "DateAndTime" Then
                        Dim bln As Boolean
                        If Boolean.TryParse(el.GetAttribute("value"), bln) Then DateAndTime = bln
                    ElseIf el.GetAttribute("key") = "X" Then
                        XField = el.GetAttribute("value")
                    ElseIf el.GetAttribute("key") = "Y" Then
                        YField = el.GetAttribute("value")
                    ElseIf el.GetAttribute("key") = "USNG" Then
                        USNGField = el.GetAttribute("value")
                    End If
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(ex.ToString)
            End Try
        End Sub
    End Class
End Namespace
