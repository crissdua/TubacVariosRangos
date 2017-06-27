Imports System.Windows.Forms

Public Class systemform

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Private esFactura As Boolean
    Private IdForm As String
    Private IdItem As String
    Private IdEvent As Integer
    Private IdAction As Boolean = False

    Private Sub SetApplication()


        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi

            '// by following the steps specified above, the following
            '// statment should be suficient for either development or run mode

            sConnectionString = Utilss.ConnectionString  'Environment.GetCommandLineArgs.GetValue(1)

            '// connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString)

            '// get an initialized application object

            SBO_Application = SboGuiApi.GetApplication()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Ocurrio un error")
        End Try

    End Sub

    Public Sub New()
        MyBase.New()
        Try

            SetApplication()
            Dim result As Integer
            Dim lerrcode As Integer
            Dim serrmsg As String = ""

            If Not SetConnectionContext() = 0 Then
                SBO_Application.MessageBox("Failed setting a connection to DI API")
                End ' Terminating the Add-On Application
            End If

            result = ConnectToCompany()
            If Not result = 0 Then
                SBO_Application.MessageBox(result & " Failed connecting to the company's Data Base")
                End ' Terminating the Add-On Application
            End If


            SBO_Application.StatusBar.SetText("Iniciando Addon Cambio de Valor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Utilss.SBOApplication = SBO_Application
            Utilss.Company = oCompany

            AddMenuItems()
            SetFilters()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message & vbNewLine & "SBO application not found")
            System.Windows.Forms.Application.Exit()
        End Try
    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String
        Dim sConnectionContext As String
        Dim lRetCode As Integer

        '// First initialize the Company object

        oCompany = New SAPbobsCOM.Company

        '// Acquire the connection context cookie from the DI API.
        sCookie = oCompany.GetContextCookie

        '// Retrieve the connection context string from the UI API using the
        '// acquired cookie.
        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

        '// before setting the SBO Login Context make sure the company is not
        '// connected

        If oCompany.Connected = True Then
            oCompany.Disconnect()
        End If

        '// Set the connection context information to the DI API.
        SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

    End Function

    Private Function ConnectToCompany() As Integer

        Try
            '// Make sure you're not already connected.
            If oCompany.Connected = True Then
                oCompany.Disconnect()
            End If

            'oCompany = SBO_Application.Company.GetDICompany

            '// Establish the connection to the company database.
            ConnectToCompany = oCompany.Connect
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try

    End Function
    Private Sub SetFilters()
        
        '// Create a new EventFilters object

        oFilters = New SAPbouiCOM.EventFilters



        '// add an event type to the container

        '// this method returns an EventFilter object

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)

        'oFilter = oFilter.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        ' oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        ' oFilter.AddEx("60006") 'Quotation Form


        oFilter.AddEx("60004")
        'oFilter.AddEx("139") 'Orders Form
        'oFilter.AddEx("133") 'Invoice Form
        'oFilter.AddEx("169") 'Main Menu
        SBO_Application.SetFilter(oFilters)

    End Sub
    Private Sub AddMenuItems()

        Try

            '//******************************************************************
            '// Let's add a separator, a pop-up menu item and a string menu item
            '//******************************************************************

            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem

            '// Get the menus collection from the application
            oMenus = SBO_Application.Menus
            '--------------------------------------------
            'Save an XML file containing the menus...
            '--------------------------------------------
            'sXML = SBO_Application.Menus.GetAsXML
            'Dim xmlD As System.Xml.XmlDocument
            'xmlD = New System.Xml.XmlDocument
            'xmlD.LoadXml(sXML)
            'xmlD.Save("c:\\mnu.xml")
            '--------------------------------------------


            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuItem = SBO_Application.Menus.Item("43520") 'modulo de menu'

            Dim sPath As String

            sPath = System.Windows.Forms.Application.StartupPath
            sPath = sPath.Remove(sPath.Length - 3, 3)

            '// find the place in wich you want to add your menu item
            '// in this example I chose to add my menu item under
            '// SAP Business One.
            If SBO_Application.Menus.Exists("CambioValor") Then
                SBO_Application.Menus.RemoveEx("CambioValor")
            End If
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "CambioValor"
            oCreationPackage.String = "Cambio de Valor"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = Replace(System.Windows.Forms.Application.StartupPath & "\valor.png", "\\", "\")
            oCreationPackage.Position = 1

            oMenus = oMenuItem.SubMenus

            Try ' If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage)

                '// Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("CambioValor")
                oMenus = oMenuItem.SubMenus

                '// Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "pantalla1"
                oCreationPackage.String = "Cambio de Valor"
                oMenus.AddEx(oCreationPackage)


            Catch er As Exception ' Menu already exists
                'SBO_Application.MessageBox("Menu Already Exists")
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

        If (pVal.MenuUID = "pantalla1") And (pVal.BeforeAction = False) Then
            Dim initFrm As New pantalla1
            BubbleEvent = False
        End If

    End Sub
End Class
