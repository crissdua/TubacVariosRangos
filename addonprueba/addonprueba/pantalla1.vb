
    Imports System.Globalization.CultureInfo
Public Class pantalla1
    Dim XmlForm As String = Replace(System.Windows.Forms.Application.StartupPath & "\pantalla1.srf", "\\", "\")

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Dim lineinioriginal As Integer
    Dim linefinoriginal As Integer
    Dim oGrid As SAPbouiCOM.Grid


    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = Utilss.SBOApplication
            Me.oCompany = Utilss.Company

            If Utilss.ActivateFormIsOpen(SBO_Application, "FrmValor") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("FrmValor")
                oForm.Left = 400
                oForm.DataSources.DataTables.Add("MyDataTable")
                Dim otro As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)
                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_PRICE)


#Region "ItemsValPorc"
                Dim oPrecioUpdate As SAPbouiCOM.EditText
                oPrecioUpdate = oForm.Items.Item("ItemVal").Specific
                oPrecioUpdate.DataBind.SetBound(True, "", "Date")
                oForm.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento As SAPbouiCOM.EditText
                oDescuento = oForm.Items.Item("ItemPorc").Specific
                oDescuento.DataBind.SetBound(True, "", "Date1")
                '2
                Dim oPrecioUpdate2 As SAPbouiCOM.EditText
                oPrecioUpdate2 = oForm.Items.Item("ItemVal2").Specific
                oPrecioUpdate2.DataBind.SetBound(True, "", "Date")
                oForm.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento2 As SAPbouiCOM.EditText
                oDescuento2 = oForm.Items.Item("ItemPorc2").Specific
                oDescuento2.DataBind.SetBound(True, "", "Date2")
                '3
                Dim oPrecioUpdate3 As SAPbouiCOM.EditText
                oPrecioUpdate3 = oForm.Items.Item("ItemVal3").Specific
                oPrecioUpdate3.DataBind.SetBound(True, "", "Date")
                oForm.DataSources.UserDataSources.Add("Date3", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento3 As SAPbouiCOM.EditText
                oDescuento3 = oForm.Items.Item("ItemPorc3").Specific
                oDescuento3.DataBind.SetBound(True, "", "Date3")
                '4
                Dim oPrecioUpdate4 As SAPbouiCOM.EditText
                oPrecioUpdate4 = oForm.Items.Item("ItemVal4").Specific
                oPrecioUpdate4.DataBind.SetBound(True, "", "Date")
                oForm.DataSources.UserDataSources.Add("Date4", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento4 As SAPbouiCOM.EditText
                oDescuento4 = oForm.Items.Item("ItemPorc4").Specific
                oDescuento4.DataBind.SetBound(True, "", "Date4")
                '5
                Dim oPrecioUpdate5 As SAPbouiCOM.EditText
                oPrecioUpdate5 = oForm.Items.Item("ItemVal5").Specific
                oPrecioUpdate5.DataBind.SetBound(True, "", "Date")
                oForm.DataSources.UserDataSources.Add("Date5", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento5 As SAPbouiCOM.EditText
                oDescuento5 = oForm.Items.Item("ItemPorc5").Specific
                oDescuento5.DataBind.SetBound(True, "", "Date5")
                '6
                Dim oPrecioUpdate6 As SAPbouiCOM.EditText
                oPrecioUpdate6 = oForm.Items.Item("ItemVal6").Specific
                oPrecioUpdate6.DataBind.SetBound(True, "", "Date")
                oForm.DataSources.UserDataSources.Add("Date6", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento6 As SAPbouiCOM.EditText
                oDescuento6 = oForm.Items.Item("ItemPorc6").Specific
                oDescuento6.DataBind.SetBound(True, "", "Date6")
                '7
                Dim oPrecioUpdate7 As SAPbouiCOM.EditText
                oPrecioUpdate7 = oForm.Items.Item("ItemVal7").Specific
                oPrecioUpdate7.DataBind.SetBound(True, "", "Date")
                oForm.DataSources.UserDataSources.Add("Date7", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento7 As SAPbouiCOM.EditText
                oDescuento7 = oForm.Items.Item("ItemPorc7").Specific
                oDescuento7.DataBind.SetBound(True, "", "Date7")
#End Region


                oGrid = oForm.Items.Item("grdDatos").Specific
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

#Region "CheckPorTon"
                Dim oChkPor As SAPbouiCOM.CheckBox
                oChkPor = oForm.Items.Item("ChkPor").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor.DataBind.SetBound(True, "", "ChkPor")
                oForm.DataSources.UserDataSources.Item("ChkPor").Value = "N"
                Dim oChkTon As SAPbouiCOM.CheckBox
                oChkTon = oForm.Items.Item("ChkTon").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon.DataBind.SetBound(True, "", "ChkTon")
                oForm.DataSources.UserDataSources.Item("ChkTon").Value = "N"
                '2
                Dim oChkPor2 As SAPbouiCOM.CheckBox
                oChkPor2 = oForm.Items.Item("ChkPor2").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor2.DataBind.SetBound(True, "", "ChkPor2")
                oForm.DataSources.UserDataSources.Item("ChkPor2").Value = "N"
                Dim oChkTon2 As SAPbouiCOM.CheckBox
                oChkTon2 = oForm.Items.Item("ChkTon2").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon2.DataBind.SetBound(True, "", "ChkTon2")
                oForm.DataSources.UserDataSources.Item("ChkTon2").Value = "N"
                '3
                Dim oChkPor3 As SAPbouiCOM.CheckBox
                oChkPor3 = oForm.Items.Item("ChkPor3").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor3.DataBind.SetBound(True, "", "ChkPor3")
                oForm.DataSources.UserDataSources.Item("ChkPor3").Value = "N"
                Dim oChkTon3 As SAPbouiCOM.CheckBox
                oChkTon3 = oForm.Items.Item("ChkTon3").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon3.DataBind.SetBound(True, "", "ChkTon3")
                oForm.DataSources.UserDataSources.Item("ChkTon3").Value = "N"
                '4
                Dim oChkPor4 As SAPbouiCOM.CheckBox
                oChkPor4 = oForm.Items.Item("ChkPor4").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor4.DataBind.SetBound(True, "", "ChkPor4")
                oForm.DataSources.UserDataSources.Item("ChkPor4").Value = "N"
                Dim oChkTon4 As SAPbouiCOM.CheckBox
                oChkTon4 = oForm.Items.Item("ChkTon4").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon4.DataBind.SetBound(True, "", "ChkTon4")
                oForm.DataSources.UserDataSources.Item("ChkTon4").Value = "N"
                '5
                Dim oChkPor5 As SAPbouiCOM.CheckBox
                oChkPor5 = oForm.Items.Item("ChkPor5").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor5.DataBind.SetBound(True, "", "ChkPor5")
                oForm.DataSources.UserDataSources.Item("ChkPor5").Value = "N"
                Dim oChkTon5 As SAPbouiCOM.CheckBox
                oChkTon5 = oForm.Items.Item("ChkTon5").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon5.DataBind.SetBound(True, "", "ChkTon5")
                oForm.DataSources.UserDataSources.Item("ChkTon5").Value = "N"
                '6
                Dim oChkPor6 As SAPbouiCOM.CheckBox
                oChkPor6 = oForm.Items.Item("ChkPor6").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor6.DataBind.SetBound(True, "", "ChkPor6")
                oForm.DataSources.UserDataSources.Item("ChkPor6").Value = "N"
                Dim oChkTon6 As SAPbouiCOM.CheckBox
                oChkTon6 = oForm.Items.Item("ChkTon6").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon6.DataBind.SetBound(True, "", "ChkTon6")
                oForm.DataSources.UserDataSources.Item("ChkTon6").Value = "N"
                '7
                Dim oChkPor7 As SAPbouiCOM.CheckBox
                oChkPor7 = oForm.Items.Item("ChkPor7").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor7.DataBind.SetBound(True, "", "ChkPor7")
                oForm.DataSources.UserDataSources.Item("ChkPor7").Value = "N"
                Dim oChkTon7 As SAPbouiCOM.CheckBox
                oChkTon7 = oForm.Items.Item("ChkTon7").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon7.DataBind.SetBound(True, "", "ChkTon7")
                oForm.DataSources.UserDataSources.Item("ChkTon7").Value = "N"
#End Region

            Else
                oForm = Me.SBO_Application.Forms.Item("FrmValor")
                oForm.Left = 400
                oForm.Visible = true
            End If

        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub

    Private Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        oXmlDoc.Load(FileName)
        SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub

    Private Sub LlenaGrid(valor As String)
        Try
            Dim QryStr As String


            QryStr = (String.Format("select Itemcode,(LineNum + 1) 'Linea', Dscription 'Descripcion', Quantity 'Cantidad', Price 'Precio', DiscPrcnt 'Descuento' from QUT1 where DocEntry = '{0}'", valor))
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(QryStr)
            oGrid = oForm.Items.Item("grdDatos").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            oGrid.Columns.GetEnumerator()
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(5).Editable = False
            CType(oGrid.Columns.Item(0), SAPbouiCOM.EditTextColumn).LinkedObjectType = 4
            linefinoriginal = oGrid.Rows.Count.ToString()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
            oGrid = Nothing
            GC.Collect()
        Catch ex As Exception
            'SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub

    Private Sub CambiaTon(doc As String, val As String, line1 As Integer, line2 As Integer)
        Try
            Dim orecord As SAPbobsCOM.Recordset
            Dim linea1 As Integer
            Dim linea2 As Integer
            Dim precio As Decimal
            Dim precio2 As Decimal
            linea1 = Convert.ToInt32(line1)
            linea2 = Convert.ToInt32(line2)
            Dim valores As string

            precio = GetDouble(val)

            valores = Replace(Convert.ToString(precio), ",", ".")

            'Dim oQuote As SAPbobsCOM.Documents
            Dim oError As Integer = -1
            Dim message As String = ""
            Dim oQuote As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)


            For a As Integer = linea1 To linea2

                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orecord.DoQuery("select (((" + valores + ")*(isnull(IWeight1,0)*QU.Quantity/1000))/Qu.Quantity) from oitm OI join QUT1 QU on OI.ItemCode = QU.ItemCode where QU.DocEntry = '" + doc + "' and LineNum = '" + a.ToString + "'")
                precio2 = Convert.ToDecimal(orecord.Fields.Item(0).Value)
                If oQuote.GetByKey(doc) Then
                    oQuote.Lines.SetCurrentLine(a)
                    oQuote.Lines.UnitPrice = precio2
                    oQuote.Lines.DiscountPercent = 0
                    oError = oQuote.Update()
                    If oError <> 0 Then
                        SBOApplication.SetStatusBarMessage("Error al actualizar Precio en Linea :" + a.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End If

                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                orecord = Nothing
                GC.Collect()
            Next
            'Dim orecord As SAPbobsCOM.Recordset
            'orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'orecord.DoQuery("Update QUT1 set Price = " + val + " where DocEntry = " + doc + " and LineNum between " + line1 + " and " + line2 + "")
            'orecord = Nothing
            'GC.Collect()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub CambiaPor(doc As String, porc As String, line1 As Integer, line2 As Integer)
        Try
            Dim linea1 As Integer
            Dim linea2 As Integer
            Dim Desc As Double
            linea1 = Convert.ToInt32(line1)
            linea2 = Convert.ToInt32(line2)

            Desc = Convert.ToDecimal(porc)

            Dim oQuote As SAPbobsCOM.Documents
            Dim oError As Integer = -1
            Dim message As String = ""
            oQuote = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)

            For a As Integer = linea1 To linea2


                If oQuote.GetByKey(doc) Then
                    oQuote.Lines.SetCurrentLine(a)
                    oQuote.Lines.DiscountPercent = Desc
                    oError = oQuote.Update()
                    If oError <> 0 Then
                        SBOApplication.SetStatusBarMessage(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End If

                End If
            Next
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        Try
            If pVal.FormUID = "FrmValor" Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBO_Application.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    If oCFLEvento.BeforeAction = False Then
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvento.SelectedObjects
                        Dim val As String


                        If (pVal.ItemUID = "Item_0") Then
                            Try
                                Dim txtFactura As SAPbouiCOM.EditText = oForm.Items.Item("Item_0").Specific
                                val = oDataTable.GetValue("DocEntry", 0)
                                LlenaGrid(val)
                                txtFactura.Value = val
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If

            End If

#Region "muestra porcentaje o valor"
#Region "ChkPor"
            If pVal.ItemUID = "ChkPor" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor").Specific
                Dim ChkTon As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon").Specific
                Dim TxtPorc As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc").Specific
                Dim Lblpor As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor").Specific
                Dim LblVal As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor").Specific
                Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific


                If ChkPor.Checked = True And ChkTon.Checked = True Then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If
                If ChkPor.Checked = True And ChkTon.Checked = False Then
                    TxtPorc.Item.Visible = True
                    Lblpor.Item.Visible = True
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If
                If ChkPor.Checked = False And ChkTon.Checked = False Then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If
                If ChkPor.Checked = False And ChkTon.Checked = True Then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = True
                    LblVal.Item.Visible = True
                    Return
                End If
            End If
            '2
            If pVal.ItemUID = "ChkPor2" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor2 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor2").Specific
                Dim ChkTon2 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon2").Specific
                Dim TxtPorc2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc2").Specific
                Dim Lblpor2 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor2").Specific
                Dim LblVal2 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor2").Specific
                Dim txtValor2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal2").Specific
                If ChkPor2.Checked = True And ChkTon2.Checked = True Then
                    TxtPorc2.Item.Visible = False
                    Lblpor2.Item.Visible = False
                    txtValor2.Item.Visible = False
                    LblVal2.Item.Visible = False
                    Return
                End If
                If ChkPor2.Checked = True And ChkTon2.Checked = False Then
                    TxtPorc2.Item.Visible = True
                    Lblpor2.Item.Visible = True
                    txtValor2.Item.Visible = False
                    LblVal2.Item.Visible = False
                    Return
                End If
                If ChkPor2.Checked = False And ChkTon2.Checked = False Then
                    TxtPorc2.Item.Visible = False
                    Lblpor2.Item.Visible = False
                    txtValor2.Item.Visible = False
                    LblVal2.Item.Visible = False
                    Return
                End If
                If ChkPor2.Checked = False And ChkTon2.Checked = True Then
                    TxtPorc2.Item.Visible = False
                    Lblpor2.Item.Visible = False
                    txtValor2.Item.Visible = True
                    LblVal2.Item.Visible = True
                    Return
                End If
            End If
            '3
            If pVal.ItemUID = "ChkPor3" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor3 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor3").Specific
                Dim ChkTon3 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon3").Specific
                Dim TxtPorc3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc3").Specific
                Dim Lblpor3 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor3").Specific
                Dim LblVal3 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor3").Specific
                Dim txtValor3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal3").Specific
                If ChkPor3.Checked = True And ChkTon3.Checked = True Then
                    TxtPorc3.Item.Visible = False
                    Lblpor3.Item.Visible = False
                    txtValor3.Item.Visible = False
                    LblVal3.Item.Visible = False
                    Return
                End If
                If ChkPor3.Checked = True And ChkTon3.Checked = False Then
                    TxtPorc3.Item.Visible = True
                    Lblpor3.Item.Visible = True
                    txtValor3.Item.Visible = False
                    LblVal3.Item.Visible = False
                    Return
                End If
                If ChkPor3.Checked = False And ChkTon3.Checked = False Then
                    TxtPorc3.Item.Visible = False
                    Lblpor3.Item.Visible = False
                    txtValor3.Item.Visible = False
                    LblVal3.Item.Visible = False
                    Return
                End If
                If ChkPor3.Checked = False And ChkTon3.Checked = True Then
                    TxtPorc3.Item.Visible = False
                    Lblpor3.Item.Visible = False
                    txtValor3.Item.Visible = True
                    LblVal3.Item.Visible = True
                    Return
                End If
            End If
            '4
            If pVal.ItemUID = "ChkPor4" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor4 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor4").Specific
                Dim ChkTon4 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon4").Specific
                Dim TxtPorc4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc4").Specific
                Dim Lblpor4 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor4").Specific
                Dim LblVal4 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor4").Specific
                Dim txtValor4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal4").Specific
                If ChkPor4.Checked = True And ChkTon4.Checked = True Then
                    TxtPorc4.Item.Visible = False
                    Lblpor4.Item.Visible = False
                    txtValor4.Item.Visible = False
                    LblVal4.Item.Visible = False
                    Return
                End If
                If ChkPor4.Checked = True And ChkTon4.Checked = False Then
                    TxtPorc4.Item.Visible = True
                    Lblpor4.Item.Visible = True
                    txtValor4.Item.Visible = False
                    LblVal4.Item.Visible = False
                    Return
                End If
                If ChkPor4.Checked = False And ChkTon4.Checked = False Then
                    TxtPorc4.Item.Visible = False
                    Lblpor4.Item.Visible = False
                    txtValor4.Item.Visible = False
                    LblVal4.Item.Visible = False
                    Return
                End If
                If ChkPor4.Checked = False And ChkTon4.Checked = True Then
                    TxtPorc4.Item.Visible = False
                    Lblpor4.Item.Visible = False
                    txtValor4.Item.Visible = True
                    LblVal4.Item.Visible = True
                    Return
                End If
            End If
            '5
            If pVal.ItemUID = "ChkPor5" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor5 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor5").Specific
                Dim ChkTon5 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon5").Specific
                Dim TxtPorc5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc5").Specific
                Dim Lblpor5 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor5").Specific
                Dim LblVal5 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor5").Specific
                Dim txtValor5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal5").Specific
                If ChkPor5.Checked = True And ChkTon5.Checked = True Then
                    TxtPorc5.Item.Visible = False
                    Lblpor5.Item.Visible = False
                    txtValor5.Item.Visible = False
                    LblVal5.Item.Visible = False
                    Return
                End If
                If ChkPor5.Checked = True And ChkTon5.Checked = False Then
                    TxtPorc5.Item.Visible = True
                    Lblpor5.Item.Visible = True
                    txtValor5.Item.Visible = False
                    LblVal5.Item.Visible = False
                    Return
                End If
                If ChkPor5.Checked = False And ChkTon5.Checked = False Then
                    TxtPorc5.Item.Visible = False
                    Lblpor5.Item.Visible = False
                    txtValor5.Item.Visible = False
                    LblVal5.Item.Visible = False
                    Return
                End If
                If ChkPor5.Checked = False And ChkTon5.Checked = True Then
                    TxtPorc5.Item.Visible = False
                    Lblpor5.Item.Visible = False
                    txtValor5.Item.Visible = True
                    LblVal5.Item.Visible = True
                    Return
                End If
            End If
            '6
            If pVal.ItemUID = "ChkPor6" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor6 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor6").Specific
                Dim ChkTon6 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon6").Specific
                Dim TxtPorc6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc6").Specific
                Dim Lblpor6 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor6").Specific
                Dim LblVal6 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor6").Specific
                Dim txtValor6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal6").Specific
                If ChkPor6.Checked = True And ChkTon6.Checked = True Then
                    TxtPorc6.Item.Visible = False
                    Lblpor6.Item.Visible = False
                    txtValor6.Item.Visible = False
                    LblVal6.Item.Visible = False
                    Return
                End If
                If ChkPor6.Checked = True And ChkTon6.Checked = False Then
                    TxtPorc6.Item.Visible = True
                    Lblpor6.Item.Visible = True
                    txtValor6.Item.Visible = False
                    LblVal6.Item.Visible = False
                    Return
                End If
                If ChkPor6.Checked = False And ChkTon6.Checked = False Then
                    TxtPorc6.Item.Visible = False
                    Lblpor6.Item.Visible = False
                    txtValor6.Item.Visible = False
                    LblVal6.Item.Visible = False
                    Return
                End If
                If ChkPor6.Checked = False And ChkTon6.Checked = True Then
                    TxtPorc6.Item.Visible = False
                    Lblpor6.Item.Visible = False
                    txtValor6.Item.Visible = True
                    LblVal6.Item.Visible = True
                    Return
                End If
            End If
            '7
            If pVal.ItemUID = "ChkPor7" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor7 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor7").Specific
                Dim ChkTon7 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon7").Specific
                Dim TxtPorc7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc7").Specific
                Dim Lblpor7 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor7").Specific
                Dim LblVal7 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor7").Specific
                Dim txtValor7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal7").Specific
                If ChkPor7.Checked = True And ChkTon7.Checked = True Then
                    TxtPorc7.Item.Visible = False
                    Lblpor7.Item.Visible = False
                    txtValor7.Item.Visible = False
                    LblVal7.Item.Visible = False
                    Return
                End If
                If ChkPor7.Checked = True And ChkTon7.Checked = False Then
                    TxtPorc7.Item.Visible = True
                    Lblpor7.Item.Visible = True
                    txtValor7.Item.Visible = False
                    LblVal7.Item.Visible = False
                    Return
                End If
                If ChkPor7.Checked = False And ChkTon7.Checked = False Then
                    TxtPorc7.Item.Visible = False
                    Lblpor7.Item.Visible = False
                    txtValor7.Item.Visible = False
                    LblVal7.Item.Visible = False
                    Return
                End If
                If ChkPor7.Checked = False And ChkTon7.Checked = True Then
                    TxtPorc7.Item.Visible = False
                    Lblpor7.Item.Visible = False
                    txtValor7.Item.Visible = True
                    LblVal7.Item.Visible = True
                    Return
                End If
            End If
#End Region
#Region "ChkTon"
            If pVal.ItemUID = "ChkTon" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                ' oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkPor As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor").Specific
                Dim ChkTon As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon").Specific
                Dim TxtPorc As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc").Specific
                Dim Lblpor As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor").Specific
                Dim LblVal As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor").Specific
                Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific
                If ChkTon.Checked = True And ChkPor.Checked = False Then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = True
                    LblVal.Item.Visible = True
                    Return
                End If
                If ChkPor.Checked = True And ChkTon.Checked = True Then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If
                If ChkTon.Checked = False And ChkPor.Checked = True Then
                    TxtPorc.Item.Visible = True
                    Lblpor.Item.Visible = True
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If

                If ChkTon.Checked = False And ChkPor.Checked = False Then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If

            End If
            '2
            If pVal.ItemUID = "ChkTon2" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                ' oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkPor2 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor2").Specific
                Dim ChkTon2 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon2").Specific
                Dim TxtPorc2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc2").Specific
                Dim Lblpor2 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor2").Specific
                Dim LblVal2 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor2").Specific
                Dim txtValor2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal2").Specific
                If ChkTon2.Checked = True And ChkPor2.Checked = False Then
                    TxtPorc2.Item.Visible = False
                    Lblpor2.Item.Visible = False
                    txtValor2.Item.Visible = True
                    LblVal2.Item.Visible = True
                    Return
                End If
                If ChkPor2.Checked = True And ChkTon2.Checked = True Then
                    TxtPorc2.Item.Visible = False
                    Lblpor2.Item.Visible = False
                    txtValor2.Item.Visible = False
                    LblVal2.Item.Visible = False
                    Return
                End If
                If ChkTon2.Checked = False And ChkPor2.Checked = True Then
                    TxtPorc2.Item.Visible = True
                    Lblpor2.Item.Visible = True
                    txtValor2.Item.Visible = False
                    LblVal2.Item.Visible = False
                    Return
                End If

                If ChkTon2.Checked = False And ChkPor2.Checked = False Then
                    TxtPorc2.Item.Visible = False
                    Lblpor2.Item.Visible = False
                    txtValor2.Item.Visible = False
                    LblVal2.Item.Visible = False
                    Return
                End If

            End If
            '3
            If pVal.ItemUID = "ChkTon3" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                'oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkPor3 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor3").Specific
                Dim ChkTon3 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon3").Specific
                Dim TxtPorc3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc3").Specific
                Dim Lblpor3 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor3").Specific
                Dim LblVal3 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor3").Specific
                Dim txtValor3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal3").Specific
                If ChkTon3.Checked = True And ChkPor3.Checked = False Then
                    TxtPorc3.Item.Visible = False
                    Lblpor3.Item.Visible = False
                    txtValor3.Item.Visible = True
                    LblVal3.Item.Visible = True
                    Return
                End If
                If ChkPor3.Checked = True And ChkTon3.Checked = True Then
                    TxtPorc3.Item.Visible = False
                    Lblpor3.Item.Visible = False
                    txtValor3.Item.Visible = False
                    LblVal3.Item.Visible = False
                    Return
                End If
                If ChkTon3.Checked = False And ChkPor3.Checked = True Then
                    TxtPorc3.Item.Visible = True
                    Lblpor3.Item.Visible = True
                    txtValor3.Item.Visible = False
                    LblVal3.Item.Visible = False
                    Return
                End If

                If ChkTon3.Checked = False And ChkPor3.Checked = False Then
                    TxtPorc3.Item.Visible = False
                    Lblpor3.Item.Visible = False
                    txtValor3.Item.Visible = False
                    LblVal3.Item.Visible = False
                    Return
                End If

            End If
            '4
            If pVal.ItemUID = "ChkTon4" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                'oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkPor4 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor4").Specific
                Dim ChkTon4 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon4").Specific
                Dim TxtPorc4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc4").Specific
                Dim Lblpor4 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor4").Specific
                Dim LblVal4 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor4").Specific
                Dim txtValor4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal4").Specific
                If ChkTon4.Checked = True And ChkPor4.Checked = False Then
                    TxtPorc4.Item.Visible = False
                    Lblpor4.Item.Visible = False
                    txtValor4.Item.Visible = True
                    LblVal4.Item.Visible = True
                    Return
                End If
                If ChkPor4.Checked = True And ChkTon4.Checked = True Then
                    TxtPorc4.Item.Visible = False
                    Lblpor4.Item.Visible = False
                    txtValor4.Item.Visible = False
                    LblVal4.Item.Visible = False
                    Return
                End If
                If ChkTon4.Checked = False And ChkPor4.Checked = True Then
                    TxtPorc4.Item.Visible = True
                    Lblpor4.Item.Visible = True
                    txtValor4.Item.Visible = False
                    LblVal4.Item.Visible = False
                    Return
                End If

                If ChkTon4.Checked = False And ChkPor4.Checked = False Then
                    TxtPorc4.Item.Visible = False
                    Lblpor4.Item.Visible = False
                    txtValor4.Item.Visible = False
                    LblVal4.Item.Visible = False
                    Return
                End If

            End If
            '5
            If pVal.ItemUID = "ChkTon5" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                'oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkPor5 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor5").Specific
                Dim ChkTon5 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon5").Specific
                Dim TxtPorc5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc5").Specific
                Dim Lblpor5 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor5").Specific
                Dim LblVal5 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor5").Specific
                Dim txtValor5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal5").Specific
                If ChkTon5.Checked = True And ChkPor5.Checked = False Then
                    TxtPorc5.Item.Visible = False
                    Lblpor5.Item.Visible = False
                    txtValor5.Item.Visible = True
                    LblVal5.Item.Visible = True
                    Return
                End If
                If ChkPor5.Checked = True And ChkTon5.Checked = True Then
                    TxtPorc5.Item.Visible = False
                    Lblpor5.Item.Visible = False
                    txtValor5.Item.Visible = False
                    LblVal5.Item.Visible = False
                    Return
                End If
                If ChkTon5.Checked = False And ChkPor5.Checked = True Then
                    TxtPorc5.Item.Visible = True
                    Lblpor5.Item.Visible = True
                    txtValor5.Item.Visible = False
                    LblVal5.Item.Visible = False
                    Return
                End If

                If ChkTon5.Checked = False And ChkPor5.Checked = False Then
                    TxtPorc5.Item.Visible = False
                    Lblpor5.Item.Visible = False
                    txtValor5.Item.Visible = False
                    LblVal5.Item.Visible = False
                    Return
                End If

            End If
            '6
            If pVal.ItemUID = "ChkTon6" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                'oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkPor6 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor6").Specific
                Dim ChkTon6 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon6").Specific
                Dim TxtPorc6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc6").Specific
                Dim Lblpor6 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor6").Specific
                Dim LblVal6 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor6").Specific
                Dim txtValor6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal6").Specific
                If ChkTon6.Checked = True And ChkPor6.Checked = False Then
                    TxtPorc6.Item.Visible = False
                    Lblpor6.Item.Visible = False
                    txtValor6.Item.Visible = True
                    LblVal6.Item.Visible = True
                    Return
                End If
                If ChkPor6.Checked = True And ChkTon6.Checked = True Then
                    TxtPorc6.Item.Visible = False
                    Lblpor6.Item.Visible = False
                    txtValor6.Item.Visible = False
                    LblVal6.Item.Visible = False
                    Return
                End If
                If ChkTon6.Checked = False And ChkPor6.Checked = True Then
                    TxtPorc6.Item.Visible = True
                    Lblpor6.Item.Visible = True
                    txtValor6.Item.Visible = False
                    LblVal6.Item.Visible = False
                    Return
                End If

                If ChkTon6.Checked = False And ChkPor6.Checked = False Then
                    TxtPorc6.Item.Visible = False
                    Lblpor6.Item.Visible = False
                    txtValor6.Item.Visible = False
                    LblVal6.Item.Visible = False
                    Return
                End If

            End If
            '7
            If pVal.ItemUID = "ChkTon7" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                ' oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkPor7 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor7").Specific
                Dim ChkTon7 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon7").Specific
                Dim TxtPorc7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc7").Specific
                Dim Lblpor7 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor7").Specific
                Dim LblVal7 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor7").Specific
                Dim txtValor7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal7").Specific
                If ChkTon7.Checked = True And ChkPor7.Checked = False Then
                    TxtPorc7.Item.Visible = False
                    Lblpor7.Item.Visible = False
                    txtValor7.Item.Visible = True
                    LblVal7.Item.Visible = True
                    Return
                End If
                If ChkPor7.Checked = True And ChkTon7.Checked = True Then
                    TxtPorc7.Item.Visible = False
                    Lblpor7.Item.Visible = False
                    txtValor7.Item.Visible = False
                    LblVal7.Item.Visible = False
                    Return
                End If
                If ChkTon7.Checked = False And ChkPor7.Checked = True Then
                    TxtPorc7.Item.Visible = True
                    Lblpor7.Item.Visible = True
                    txtValor7.Item.Visible = False
                    LblVal7.Item.Visible = False
                    Return
                End If

                If ChkTon7.Checked = False And ChkPor7.Checked = False Then
                    TxtPorc7.Item.Visible = False
                    Lblpor7.Item.Visible = False
                    txtValor7.Item.Visible = False
                    LblVal7.Item.Visible = False
                    Return
                End If

            End If
#End Region
#End Region

            If pVal.ItemUID = "cmdOk" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
#Region "Variables"
                Dim ChkPor As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor").Specific
                Dim ChkTon As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon").Specific
                Dim TxtPorc As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc").Specific
                Dim Lblpor As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor").Specific
                Dim LblVal As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor").Specific
                Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific
                '2
                Dim ChkPor2 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor2").Specific
                Dim ChkTon2 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon2").Specific
                Dim TxtPorc2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc2").Specific
                Dim Lblpor2 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor2").Specific
                Dim LblVal2 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor2").Specific
                Dim txtValor2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal2").Specific
                '3
                Dim ChkPor3 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor3").Specific
                Dim ChkTon3 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon3").Specific
                Dim TxtPorc3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc3").Specific
                Dim Lblpor3 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor3").Specific
                Dim LblVal3 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor3").Specific
                Dim txtValor3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal3").Specific
                '4
                Dim ChkPor4 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor4").Specific
                Dim ChkTon4 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon4").Specific
                Dim TxtPorc4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc4").Specific
                Dim Lblpor4 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor4").Specific
                Dim LblVal4 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor4").Specific
                Dim txtValor4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal4").Specific
                '5
                Dim ChkPor5 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor5").Specific
                Dim ChkTon5 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon5").Specific
                Dim TxtPorc5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc5").Specific
                Dim Lblpor5 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor5").Specific
                Dim LblVal5 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor5").Specific
                Dim txtValor5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal5").Specific
                '6
                Dim ChkPor6 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor6").Specific
                Dim ChkTon6 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon6").Specific
                Dim TxtPorc6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc6").Specific
                Dim Lblpor6 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor6").Specific
                Dim LblVal6 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor6").Specific
                Dim txtValor6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal6").Specific
                '7
                Dim ChkPor7 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor7").Specific
                Dim ChkTon7 As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon7").Specific
                Dim TxtPorc7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc7").Specific
                Dim Lblpor7 As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor7").Specific
                Dim LblVal7 As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor7").Specific
                Dim txtValor7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal7").Specific
#End Region

                Dim txtDocum As SAPbouiCOM.EditText = oForm.Items.Item("Item_0").Specific
                Dim TxtLineini As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni").Specific
                Dim TxtLineFin As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin").Specific
                Dim lineini As Integer
                Dim linefin As Integer
                Dim Porce As Double
                Dim Docnum As String
                Dim Valor As Double

                Docnum = txtDocum.Value.Trim


                Dim TxtLineini2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni2").Specific
                Dim TxtLineFin2 As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin2").Specific
                Dim lineini2 As Integer
                Dim linefin2 As Integer
                Dim Valor2 As Double
                Dim Porce2 As Double
                ''3
                Dim TxtLineini3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni3").Specific
                Dim TxtLineFin3 As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin3").Specific
                Dim lineini3 As Integer
                Dim linefin3 As Integer
                Dim Valor3 As Double
                Dim Porce3 As Double
                ''4
                Dim TxtLineini4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni4").Specific
                Dim TxtLineFin4 As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin4").Specific
                Dim lineini4 As Integer
                Dim linefin4 As Integer
                Dim Valor4 As Double
                Dim Porce4 As Double
                ''5
                Dim TxtLineini5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni5").Specific
                Dim TxtLineFin5 As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin5").Specific
                Dim lineini5 As Integer
                Dim linefin5 As Integer
                Dim Valor5 As Double
                Dim Porce5 As Double
                ''6
                Dim TxtLineini6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni6").Specific
                Dim TxtLineFin6 As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin6").Specific
                Dim lineini6 As Integer
                Dim linefin6 As Integer
                Dim Valor6 As Double
                Dim Porce6 As Double
                ''7
                Dim TxtLineini7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni7").Specific
                Dim TxtLineFin7 As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin7").Specific
                Dim lineini7 As Integer
                Dim linefin7 As Integer
                Dim Valor7 As Double
                Dim Porce7 As Double
#Region "valida campos en blanco"
                If txtDocum.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return
                End If
#End Region



                Dim resp = SBO_Application.MessageBox("Guardara los cambios en la Oferta de Venta con NO." & txtDocum.Value.Trim, 1, "SI", "NO")
                If resp = 1 Then
#Region "ChkTon"
                    If ChkTon.Checked = True Then



                        If txtValor.Value = "" Or txtValor.Value <= 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 1", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Valor = GetDouble(txtValor.Value)
                        lineini = (TxtLineini.Value) - 1
                        linefin = (TxtLineFin.Value) - 1
                        If lineini >= 0 And TxtLineFin.Value <= linefinoriginal And (TxtLineini.Value <= TxtLineFin.Value) Then
                            CambiaTon(Docnum, Valor, lineini, linefin)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 1 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkTon2"
                    If ChkTon2.Checked = True Then

                        If txtValor2.Value = "" Or txtValor2.Value <= 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 2", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Valor2 = GetDouble(txtValor2.Value)
                        lineini2 = (TxtLineini2.Value) - 1
                        linefin2 = (TxtLineFin2.Value) - 1
                        If lineini2 >= 0 And TxtLineFin2.Value <= linefinoriginal And (TxtLineini2.Value <= TxtLineFin2.Value) Then

                            CambiaTon(Docnum, Valor2, lineini2, linefin2)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 2 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkTon3"
                    If ChkTon3.Checked = True Then



                        If txtValor3.Value = "" Or txtValor3.Value <= 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 3", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Valor3 = GetDouble(txtValor3.Value)
                        lineini3 = (TxtLineini3.Value) - 1
                        linefin3 = (TxtLineFin3.Value) - 1
                        If lineini3 >= 0 And TxtLineFin3.Value <= linefinoriginal And (TxtLineini3.Value <= TxtLineFin3.Value) Then
                            CambiaTon(Docnum, Valor3, lineini3, linefin3)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 3 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkTon4"
                    If ChkTon4.Checked = True Then



                        If txtValor4.Value = "" Or txtValor4.Value <= 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 4", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Valor4 = GetDouble(txtValor4.Value)
                        lineini4 = (TxtLineini4.Value) - 1
                        linefin4 = (TxtLineFin4.Value) - 1
                        If lineini4 >= 0 And TxtLineFin4.Value <= linefinoriginal And (TxtLineini4.Value <= TxtLineFin4.Value) Then
                            CambiaTon(Docnum, Valor4, lineini4, linefin4)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 4 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkTon5"
                    If ChkTon5.Checked = True Then



                        If txtValor5.Value = "" Or txtValor5.Value <= 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 5", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Valor5 = GetDouble(txtValor5.Value)
                        lineini5 = (TxtLineini5.Value) - 1
                        linefin5 = (TxtLineFin5.Value) - 1
                        If lineini5 >= 0 And TxtLineFin5.Value <= linefinoriginal And (TxtLineini5.Value <= TxtLineFin5.Value) Then
                            CambiaTon(Docnum, Valor5, lineini5, linefin5)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 5 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkTon6"
                    If ChkTon6.Checked = True Then



                        If txtValor6.Value = "" Or txtValor6.Value <= 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 6", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Valor6 = GetDouble(txtValor6.Value)
                        lineini6 = (TxtLineini6.Value) - 1
                        linefin6 = (TxtLineFin6.Value) - 1
                        If lineini6 >= 0 And TxtLineFin6.Value <= linefinoriginal And (TxtLineini6.Value <= TxtLineFin6.Value) Then
                            CambiaTon(Docnum, Valor6, lineini6, linefin6)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 6 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkTon7"
                    If ChkTon7.Checked = True Then



                        If txtValor7.Value = "" Or txtValor7.Value <= 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 7", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Valor7 = GetDouble(txtValor7.Value)
                        lineini7 = (TxtLineini7.Value) - 1
                        linefin7 = (TxtLineFin7.Value) - 1
                        If lineini7 >= 0 And TxtLineFin7.Value <= linefinoriginal And (TxtLineini7.Value <= TxtLineFin7.Value) Then
                            CambiaTon(Docnum, Valor7, lineini7, linefin7)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 7 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region

#Region "ChkPor"
                    If ChkPor.Checked = True Then
                        'Porce = GetDouble(TxtPorc.Value)
                        If TxtPorc.Value = "" Or TxtPorc.Value < 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 1", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Porce = GetDouble(TxtPorc.Value)
                        lineini = (TxtLineini.Value) - 1
                        linefin = (TxtLineFin.Value) - 1
                        If (lineini >= 0 And TxtLineFin.Value <= linefinoriginal) And (TxtLineini.Value <= TxtLineFin.Value) Then
                            CambiaPor(Docnum, Porce, lineini, linefin)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 1 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkPor2"
                    If ChkPor2.Checked = True Then
                        'Porce2 = GetDouble(TxtPorc2.Value)
                        If TxtPorc2.Value = "" Or TxtPorc2.Value < 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 2", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Porce2 = GetDouble(TxtPorc2.Value)
                        lineini2 = (TxtLineini2.Value) - 1
                        linefin2 = (TxtLineFin2.Value) - 1
                        If (lineini2 >= 0 And TxtLineFin2.Value <= linefinoriginal) And (TxtLineini2.Value <= TxtLineFin2.Value) Then
                            CambiaPor(Docnum, Porce2, lineini2, linefin2)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 2 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkPor3"
                    If ChkPor3.Checked = True Then
                        'Porce3 = GetDouble(TxtPorc3.Value)
                        If TxtPorc3.Value = "" Or TxtPorc3.Value < 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 3", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Porce3 = GetDouble(TxtPorc3.Value)
                        lineini3 = (TxtLineini3.Value) - 1
                        linefin3 = (TxtLineFin3.Value) - 1
                        If (lineini3 >= 0 And TxtLineFin.Value <= linefinoriginal) And (TxtLineini3.Value <= TxtLineFin3.Value) Then
                            CambiaPor(Docnum, Porce3, lineini3, linefin3)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 3 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkPor4"
                    If ChkPor4.Checked = True Then
                        'Porce4 = GetDouble(TxtPorc4.Value)
                        If TxtPorc4.Value = "" Or TxtPorc4.Value < 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 4", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Porce4 = GetDouble(TxtPorc4.Value)
                        lineini4 = (TxtLineini4.Value) - 1
                        linefin4 = (TxtLineFin4.Value) - 1
                        If (lineini4 >= 0 And TxtLineFin4.Value <= linefinoriginal) And (TxtLineini4.Value <= TxtLineFin4.Value) Then
                            CambiaPor(Docnum, Porce4, lineini4, linefin4)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 4 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkPor5"
                    If ChkPor5.Checked = True Then
                        'Porce5 = GetDouble(TxtPorc5.Value)
                        If TxtPorc5.Value = "" Or TxtPorc5.Value < 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 5", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Porce5 = GetDouble(TxtPorc5.Value)
                        lineini5 = (TxtLineini5.Value) - 1
                        linefin5 = (TxtLineFin5.Value) - 1
                        If (lineini5 >= 0 And TxtLineFin5.Value <= linefinoriginal) And (TxtLineini5.Value <= TxtLineFin5.Value) Then
                            CambiaPor(Docnum, Porce5, lineini5, linefin5)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 5 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkPor6"
                    If ChkPor6.Checked = True Then
                        'Porce6 = GetDouble(TxtPorc6.Value)
                        If TxtPorc6.Value = "" Or TxtPorc6.Value < 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 6", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Porce6 = GetDouble(TxtPorc6.Value)
                        lineini6 = (TxtLineini6.Value) - 1
                        linefin6 = (TxtLineFin6.Value) - 1
                        If (lineini6 >= 0 And TxtLineFin6.Value <= linefinoriginal) And (TxtLineini6.Value <= TxtLineFin6.Value) Then
                            CambiaPor(Docnum, Porce6, lineini6, linefin6)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 6 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
#Region "ChkPor7"
                    If ChkPor7.Checked = True Then
                        'Porce7 = GetDouble(TxtPorc7.Value)
                        If TxtPorc7.Value = "" Or TxtPorc7.Value < 0 Then
                            SBO_Application.SetStatusBarMessage("Debe Verificar Valores en Rango 7", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return
                        End If
                        Porce7 = GetDouble(TxtPorc7.Value)
                        lineini7 = (TxtLineini7.Value) - 1
                        linefin7 = (TxtLineFin7.Value) - 1
                        If (lineini7 >= 0 And TxtLineFin7.Value <= linefinoriginal) And (TxtLineini7.Value <= TxtLineFin7.Value) Then
                            CambiaPor(Docnum, Porce7, lineini7, linefin7)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        Else
                            SBO_Application.SetStatusBarMessage("En el Rango 7 el numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                            BubbleEvent = False
                            Return
                        End If
                    End If
#End Region
                    BubbleEvent = False
                    Return
                End If
            End If

            'If pVal.ItemUID = "grdDatos" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
            '   Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific
            '    Dim val As string

            '    oGrid = oForm.Items.Item("grdDatos").Specific
            '    val = oGrid.DataTable.GetValue(1,oGrid.GetDataTableRowIndex(pVal.Row)).ToString
            '    SBO_Application.SetStatusBarMessage(val, SAPbouiCOM.BoMessageTime.bmt_Medium, True)

            'End If

            If pVal.ItemUID = "btnCancel" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Try
                    oForm.Close
                    BubbleEvent = False
                    Return
                Catch ex As Exception

                End Try
                
            End If
        Catch ex As Exception
            'SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            BubbleEvent = False
            Return
        End Try
    End Sub
    Public Shared Function GetDouble(ByVal doublestring As String) As Double
        Dim retval As Decimal
        Dim sep As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator

        Double.TryParse(Replace(Replace(doublestring, ".", sep), ",", sep), retval)
        Return retval
    End Function
End Class
