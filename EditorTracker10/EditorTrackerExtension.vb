Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports System.IO
Imports ESRI.ArcGIS.Geometry

'''<summary>
'''EditorTrackerExtension class implementing custom ESRI Editor Extension functionalities.
'''</summary>
Public Class EditorTrackerExtension
    Inherits ESRI.ArcGIS.Desktop.AddIns.Extension

    Private WithEvents EditEvents As IEditEvents_Event

    Public Sub New()

    End Sub

    Protected Overrides Sub OnStartup()
        EditEvents = My.ArcMap.Editor
        Dim di As New DirectoryInfo(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData)
        If Not di.Exists Then
            di.Create()
        End If
        Dim fi As New FileInfo(System.IO.Path.Combine(di.FullName, "EditorTracker.config"))
        If Not fi.Exists Then
            'Create the config file...FAIL!
            My.Settings.SaveTo(fi.FullName)
        Else
            My.Settings.LoadFrom(fi.FullName)
        End If
    End Sub

    Protected Overrides Sub OnShutdown()
        EditEvents = Nothing
    End Sub

    Private Sub EditEvents_OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Handles EditEvents.OnCreateFeature
        Dim pRow As IRow = TryCast(obj, IRow)
        If pRow Is Nothing Then Exit Sub
        Dim pDataset As IDataset = pRow.Table
        Dim idx As Integer = GetFieldIndex(pRow, My.Settings.EditorNameField)
        If idx >= 0 Then
            'if the editor_name field exists update it with the current editor name...
            pRow.Value(idx) = GetUser(pDataset.Workspace)
        End If
        idx = GetFieldIndex(pRow, My.Settings.EditDateField)
        If idx >= 0 Then
            'if the editor_date field exists update it with the current date and time...
            If My.Settings.DateAndTime Then
                pRow.Value(idx) = Now
            Else
                pRow.Value(idx) = Today
            End If
        End If
        UpdateXY(pRow)
    End Sub
	
	Private Function GetFieldIndex (ByVal pRow as IRow, FieldName as String) As Integer 
		Dim idx As Integer = pRow.Fields.FindField(FieldName)
        If idx > -1 OrElse FieldName.Length < 11 Then Return idx
		'Try truncated name
		idx = pRow.Fields.FindField(FieldName.Substring(0, 10))
		Return idx
	End Function

    Private Sub EditEvents_OnChangeFeature(ByVal obj As IObject) Handles EditEvents.OnChangeFeature
        Dim pRow As IRow = TryCast(obj, IRow)
        If pRow Is Nothing Then Exit Sub
        Dim pDataset As IDataset = pRow.Table
        Dim idx As Integer = GetFieldIndex(pRow, My.Settings.EditorNameField)
        If idx >= 0 Then
            'if the editor_name field exists update it with the current editor name...
            pRow.Value(idx) = GetUser(pDataset.Workspace)
        End If
        idx = GetFieldIndex(pRow, My.Settings.EditDateField)
        If idx >= 0 Then
            'if the editor_date field exists update it with the current date and time...
            If My.Settings.DateAndTime Then
                pRow.Value(idx) = Now
            Else
                pRow.Value(idx) = Today
            End If
        End If
        UpdateXY(pRow)
    End Sub

    Private Sub EditEvents_OnDeleteFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Handles EditEvents.OnDeleteFeature
        'Why is this here? How could you record who deleted the row if the row is deleted?
        Dim pRow As IRow = TryCast(obj, IRow)
        If pRow Is Nothing Then Exit Sub
        Dim pDataset As IDataset = pRow.Table
        Dim idx As Integer = GetFieldIndex(pRow, My.Settings.EditorNameField)
        If idx >= 0 Then
            'if the editor_name field exists update it with the current editor name...
            pRow.Value(idx) = GetUser(pDataset.Workspace)
        End If
        idx = GetFieldIndex(pRow, My.Settings.EditDateField)
        If idx >= 0 Then
            'if the editor_date field exists update it with the current date and time...
            If My.Settings.DateAndTime Then
                pRow.Value(idx) = Now
            Else
                pRow.Value(idx) = Today
            End If
        End If
        UpdateXY(pRow)
    End Sub

    Private Function GetUser(ByVal pWork As IWorkspace) As String
        'get the user based on settings OS user or database user
        Try
            If My.Settings.TrackDatabaseUser AndAlso pWork.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                Dim pDBConInfo As IDatabaseConnectionInfo = pWork
                Return pDBConInfo.ConnectedUser
            End If
            'For some reason, My.User.Name is an empty string...
            Return Environ("USERDOMAIN") & "\" & Environ("USERNAME")
        Catch ex As Exception
            Return Environ("USERDOMAIN") & "\" & Environ("USERNAME")
        End Try
    End Function

    Private Sub UpdateXY(ByVal obj As IObject)
        Dim pFeat As IFeature = TryCast(obj, IFeature)
        If pFeat Is Nothing OrElse pFeat.Shape.IsEmpty OrElse pFeat.Shape.GeometryType <> esriGeometryType.esriGeometryPoint Then Return
        Dim idxX As Integer = GetFieldIndex(pFeat, My.Settings.XField)
        Dim idxY As Integer = GetFieldIndex(pFeat, My.Settings.YField)
        If idxX >= 0 AndAlso idxY >= 0 Then
            Dim pPoint As IPoint = pFeat.ShapeCopy
            Dim pSpatialRefFactory As ISpatialReferenceFactory
            pSpatialRefFactory = New SpatialReferenceEnvironment
            Dim pGCS As IGeographicCoordinateSystem
            pGCS = pSpatialRefFactory.CreateGeographicCoordinateSystem(esriSRGeoCSType.esriSRGeoCS_WGS1984)
            pPoint.Project(pGCS)
            If Not pPoint Is Nothing Then
                pFeat.Value(idxX) = pPoint.X
                pFeat.Value(idxY) = pPoint.Y
                Dim idxUSNG As Integer = GetFieldIndex(pFeat, My.Settings.USNGField)
                If idxUSNG >= 0 Then
                    Dim convertPnt As IConversionMGRS = TryCast(pPoint, IConversionMGRS)
                    Dim sUSNG As String
                    sUSNG = convertPnt.CreateMGRS(8, True, esriMGRSModeEnum.esriMGRSMode_USNG)
                    sUSNG = Mid(sUSNG, 1, 3) & " " & Mid(sUSNG, 4, 2) & " " & Mid(sUSNG, 6, 5) & " " & Mid(sUSNG, 14, 5)
                    pFeat.Value(idxUSNG) = sUSNG
                End If
            End If
        End If
    End Sub
End Class
