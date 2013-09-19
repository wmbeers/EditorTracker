Imports System.Runtime.InteropServices
Imports System.Xml
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Configuration
Imports System.Collections.Specialized
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<ComClass(EditorTrackerClass.ClassId, EditorTrackerClass.InterfaceId, EditorTrackerClass.EventsId), _
 ProgId("EditorTracker.EditorTrackerClass")> _
Public Class EditorTrackerClass
    Implements ESRI.ArcGIS.esriSystem.IExtension
    Private WithEvents mEV As Editor
    Private mApp As IApplication

    Private ENF As String 'Field to store editor name
    Private EDF As String 'Field to store edit date
    Private TDU As Boolean 'Track database user. If true the database user name 
    'is retrieved from the connection properties adn stored in the ENF. If false
    'the OS username is retrieved from the USERNAME environment variable and stored in the ENF
    Private DT As Boolean 'Date and time. If true the date and time (5/6/2005 02:34:42 PM) are 
    'stored in the EDF field. If false the date (5/6/2005) is stored in the EDF.
    '**** storing the date and time can lead into conflicts on a multi user versioned editing environment.
    Private XF As String
    Private YF As String
    Private USNGF As String

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "04e459f4-8114-4d69-b842-41fba4377131"
    Public Const InterfaceId As String = "df362b6b-0e15-4eb1-adca-d5f97402fdc2"
    Public Const EventsId As String = "9bc261c7-b623-4b4f-8b50-8aca734f1ea3"
#End Region
#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Unregister(regKey)

    End Sub

#End Region
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "EditorTrackerClass"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        mEV = Nothing
        mApp = Nothing
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        mApp = CType(initializationData, IApplication)
        If mApp Is Nothing Then Return
        Dim pID As New UID
        pID.Value = "esriEditor.Editor"
        Dim pEditor As IEditor
        pEditor = mApp.FindExtensionByCLSID(pID)
        mEV = pEditor
        GetConfigSettings()
    End Sub
    Private Sub EditEvents_OnChangeFeature(ByVal obj As IObject) Handles mEV.OnChangeFeature
        'updte editor_name and edit_date when a record is changed
        If TypeOf obj Is IRow Then
            Dim pRow As IRow
            pRow = obj
            Dim pDataset As IDataset
            pDataset = pRow.Table
            If pRow.Fields.FindField(ENF) >= 0 Then
                'if the editor_name field exists update it with the current editor name...
                pRow.Value(pRow.Fields.FindField(ENF)) = GetUser(pDataset.Workspace)
            End If
            If pRow.Fields.FindField(EDF) >= 0 Then
                'if the editor_date field exists update it with the current date and time...
                If DT = False Then
                    pRow.Value(pRow.Fields.FindField(EDF)) = Microsoft.VisualBasic.Month(Now) & "/" & Microsoft.VisualBasic.Day(Now) & "/" & Microsoft.VisualBasic.Year(Now)
                Else
                    pRow.Value(pRow.Fields.FindField(EDF)) = Now
                End If

            End If
        End If
        If Not obj Is Nothing Then UpdateXY(obj)
    End Sub

    Private Sub EditEvents_OnCreateFeature(ByVal obj As IObject) Handles mEV.OnCreateFeature
        'updte editor_name and edit_date when a record is created
        If TypeOf obj Is IRow Then
            Dim pRow As IRow
            pRow = obj
            Dim pDataset As IDataset
            pDataset = pRow.Table
            If pRow.Fields.FindField(ENF) >= 0 Then
                'if the editor_name field exists update it with the current editor name...
                pRow.Value(pRow.Fields.FindField(ENF)) = GetUser(pDataset.Workspace)
            End If
            If pRow.Fields.FindField(EDF) >= 0 Then
                'if the editor_date field exists update it with the current date and time...
                If DT = False Then
                    pRow.Value(pRow.Fields.FindField(EDF)) = Microsoft.VisualBasic.Month(Now) & "/" & Microsoft.VisualBasic.Day(Now) & "/" & Microsoft.VisualBasic.Year(Now)
                Else
                    pRow.Value(pRow.Fields.FindField(EDF)) = Now
                End If
            End If
        End If
        If Not obj Is Nothing Then UpdateXY(obj)
    End Sub

    Private Sub EditEvents_OnDeleteFeature(ByVal obj As IObject) Handles mEV.OnDeleteFeature
        'updte editor_name and edit_date when a record is deleted
        If TypeOf obj Is IRow Then
            Dim pRow As IRow
            pRow = obj
            Dim pDataset As IDataset
            pDataset = pRow.Table
            If pRow.Fields.FindField(ENF) >= 0 Then
                'if the editor_name field exists update it with the current editor name...
                pRow.Value(pRow.Fields.FindField(ENF)) = GetUser(pDataset.Workspace)
            End If
            If pRow.Fields.FindField(EDF) >= 0 Then
                'if the editor_date field exists update it with the current date and time...
                If DT = False Then
                    pRow.Value(pRow.Fields.FindField(EDF)) = Microsoft.VisualBasic.Month(Now) & "/" & Microsoft.VisualBasic.Day(Now) & "/" & Microsoft.VisualBasic.Year(Now)
                Else
                    pRow.Value(pRow.Fields.FindField(EDF)) = Now
                End If
            End If
        End If
        If Not obj Is Nothing Then UpdateXY(obj)
    End Sub

    Private Function GetUser(ByVal pWork As IWorkspace) As String
        'get the user based on settings OS user or database user
        Dim pUser As String = Nothing
        Try
            If TDU = False Then
                pUser = Environ("USERDOMAIN") & "\" & Environ("USERNAME")
            ElseIf TDU = True And pWork.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                Dim pDBConInfo As IDatabaseConnectionInfo
                pDBConInfo = pWork
                pUser = pDBConInfo.ConnectedUser
            Else
                pUser = Environ("USERDOMAIN") & "\" & Environ("USERNAME")
            End If
        Catch ex As Exception
            pUser = Environ("USERDOMAIN") & "\" & Environ("USERNAME")
        End Try
        Return pUser
    End Function
    Private Sub GetConfigSettings()
        Dim ConfigFileName As String = My.Application.Info.DirectoryPath & "\EditorTracker.config"
        If File.Exists(ConfigFileName) Then
            Dim ConfigDoc As XmlDocument = New XmlDocument()
            ConfigDoc.Load(ConfigFileName)
            Dim xNode As XmlNode = ConfigDoc.GetElementsByTagName("EditorTrackerSettings").Item(0)
            Dim csh As IConfigurationSectionHandler = New NameValueSectionHandler()
            Dim nvc As NameValueCollection = CType(csh.Create(Nothing, Nothing, xNode), NameValueCollection)
            Try
                ENF = nvc("EDITORNAMEFIELD").ToString
            Catch ex As Exception
                ENF = "EDITOR_NAME"
            End Try
            Try
                EDF = nvc("EDITDATEFIELD").ToString
            Catch ex As Exception
                EDF = "EDIT_DATE"
            End Try
            Try
                TDU = Convert.ToBoolean(nvc("TRACKDATABASEUSER"))
            Catch ex As Exception
                TDU = True
            End Try
            Try
                DT = Convert.ToBoolean(nvc("DATEANDTIME"))
            Catch ex As Exception
                DT = False
            End Try
            Try
                XF = nvc("X").ToString
            Catch ex As Exception
                XF = "X"
            End Try
            Try
                YF = nvc("Y").ToString
            Catch ex As Exception
                YF = "Y"
            End Try
            Try
                USNGF = nvc("USNG").ToString
            Catch ex As Exception
                USNGF = "USNG"
            End Try
        Else
            ENF = "EDITOR_NAME"
            EDF = "EDIT_DATE"
            TDU = True
            DT = False
            XF = "X"
            YF = "Y"
        End If
    End Sub

    Private Sub UpdateXY(ByVal obj As IObject)
        Dim pObjClass As IObjectClass = obj.Class
        If TypeOf pObjClass Is FeatureClass Then
            Dim pFeat As IFeature = obj
            If pFeat.Fields.FindField(XF) >= 0 And pFeat.Fields.FindField(YF) >= 0 Then
                If Not pFeat.Shape.IsEmpty = True Then
                    If pFeat.Shape.GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                        Dim pPoint As ESRI.ArcGIS.Geometry.IPoint = pFeat.ShapeCopy
                        Dim pSpatialRefFactory As ISpatialReferenceFactory
                        pSpatialRefFactory = New SpatialReferenceEnvironment


                        Dim pGCS As IGeographicCoordinateSystem
                        pGCS = pSpatialRefFactory.CreateGeographicCoordinateSystem(esriSRGeoCSType.esriSRGeoCS_WGS1984)
                        pPoint.Project(pGCS)

                        If Not pPoint Is Nothing Then
                            pFeat.Value(pFeat.Fields.FindField(XF)) = pPoint.X
                            pFeat.Value(pFeat.Fields.FindField(YF)) = pPoint.Y

                            Dim convertPnt As IConversionMGRS = TryCast(pPoint, IConversionMGRS)
                            Dim sUSNG As String
                            sUSNG = convertPnt.CreateMGRS(8, True, esriMGRSModeEnum.esriMGRSMode_USNG)
                            sUSNG = Mid(sUSNG, 1, 3) & " " & Mid(sUSNG, 4, 2) & " " & Mid(sUSNG, 6, 5) & " " & Mid(sUSNG, 14, 5)
                            pFeat.Value(pFeat.Fields.FindField(USNGF)) = sUSNG

                        End If
                    End If
                End If
            End If
        End If
    End Sub

End Class


