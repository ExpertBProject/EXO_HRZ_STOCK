Imports SAPbouiCOM
Imports System.Xml
Public Class EXO_STOCK
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        cargamenu()
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        Path = objGlobal.refDi.OGEN.pathGeneral.ToString.Trim
        If objGlobal.SBOApp.Menus.Exists("EXO-MnCEART") = True Then
            Path &= "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnCEART.png") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnCEART").Image = Path & "\MnCEART.png"
                End If
            End If
        End If
    End Sub

    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnCEART"
                        'Cargamos pantalla de gestión.
                        If CargarFormEArt() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormEArt() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim Path As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormEArt = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CEART.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
#Region "CargaCombos"
            'Cargamos combo familia
            sSQL = " SELECT * FROM ("
            sSQL &= " SELECT '-',CAST(' ' as nVarchar(100)) ""ItmsGrpNam"" FROM ""DUMMY""  "
            sSQL &= "  UNION ALL "
            sSQL &= " SELECT CAST(""ItmsGrpCod"" as nvarchar),CAST(""ItmsGrpNam"" as nvarchar(100)) FROM ""OITB""  "
            sSQL &= " )t ORDER BY t.""ItmsGrpNam"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbFAM").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            'Cargamos combo subfamilia
            sSQL = " SELECT * FROM ("
            sSQL &= " SELECT CAST('-' as nVarchar(50)),CAST(' ' as nVarchar(100)) ""U_EXO_DESSUBFAM"" FROM ""DUMMY""  "
            sSQL &= "  UNION ALL "
            sSQL &= " SELECT DISTINCT ""U_EXO_CODSUBFAM"",""U_EXO_DESSUBFAM"" FROM ""@EXO_FAMSUBFAM"" "
            sSQL &= " )t ORDER BY t.""U_EXO_DESSUBFAM"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbSFAM").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

            'Cargamos Almacenes
            sSQL = " SELECT * FROM ("
            sSQL &= " SELECT CAST('-' as nVarchar(50)),CAST(' ' as nVarchar(100)) ""WhsName"" FROM ""DUMMY""  "
            sSQL &= "  UNION ALL "
            sSQL &= " SELECT DISTINCT ""WhsCode"",""WhsName"" FROM ""OWHS"" "
            sSQL &= " )t ORDER BY t.""WhsName"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
#End Region
            CargarFormEArt = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CEART"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                    If EventHandler_DOUBLE_CLICK_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CEART"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CEART"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CEART"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing


        EventHandler_Choose_FromList_After = False
        Dim sChoosefromlist As String = ""

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Dim sCod As String = "" : Dim sDes As String = ""
            If pVal.ItemUID = "txtICode" Then
                Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "4"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("ItemCode", 0).ToString

                                    oForm.DataSources.UserDataSources.Item("UDICODE").Value = sCod
                                Catch ex As Exception
                                    oForm.DataSources.UserDataSources.Item("UDICODE").Value = sCod
                                End Try
                        End Select
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
            ElseIf pVal.ItemUID = "txtIC" Then
                Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "2"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("CardCode", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("CardName", 0).ToString

                                    oForm.DataSources.UserDataSources.Item("UDIC").Value = sCod
                                    oForm.DataSources.UserDataSources.Item("UDICNAME").Value = sDes
                                Catch ex As Exception
                                    oForm.DataSources.UserDataSources.Item("UDIC").Value = sCod
                                    oForm.DataSources.UserDataSources.Item("UDICNAME").Value = sDes
                                End Try
                        End Select
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean

        EventHandler_COMBO_SELECT_After = False
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sSQL As String = "" : Dim sFam As String = ""
        Try
            oForm.Freeze(True)
            'Cargamos combo subfamilia
            If pVal.ItemUID = "cbFAM" Then
                sFam = CType(oForm.Items.Item("cbFAM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                If sFam.Trim = "-" Or sFam.Trim = "" Then
                    sSQL = "SELECT * FROM ("
                    sSQL &= " SELECT CAST('-' as nVarchar(50)),CAST(' ' as nVarchar(100)) ""U_EXO_DESSUBFAM"" FROM ""DUMMY""  "
                    sSQL &= "  UNION ALL "
                    sSQL &= "SELECT DISTINCT ""U_EXO_CODSUBFAM"",""U_EXO_DESSUBFAM"" FROM ""@EXO_FAMSUBFAM"" "
                    sSQL &= " )t ORDER BY t.""U_EXO_DESSUBFAM"" "
                Else
                    sSQL = "SELECT * FROM ("
                    sSQL &= " SELECT CAST('-' as nVarchar(50)),CAST(' ' as nVarchar(100)) ""U_EXO_DESSUBFAM"" FROM ""DUMMY""  "
                    sSQL &= "  UNION ALL "
                    sSQL &= "SELECT DISTINCT ""U_EXO_CODSUBFAM"",""U_EXO_DESSUBFAM"" FROM ""@EXO_FAMSUBFAM"" WHERE  ""U_EXO_CODFAM""='" & sFam & "' "
                    sSQL &= " )t ORDER BY t.""U_EXO_DESSUBFAM"" "
                End If
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbSFAM").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            End If


            EventHandler_COMBO_SELECT_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sSQL As String = ""
        Dim sFecha As String = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
        Dim sCodArt As String = "" : Dim sDesArt As String = ""
        Dim sFamilia As String = "" : Dim sSFamilia As String = ""
        Dim sMarca As String = ""
        Dim sAlmacen As String = ""
        Dim sIC As String = ""
        EventHandler_ItemPressed_After = False

        Try
            sCodArt = oForm.DataSources.UserDataSources.Item("UDICODE").Value.ToString
            sDesArt = oForm.DataSources.UserDataSources.Item("UDINAME").Value.ToString
            If CType(oForm.Items.Item("cbFAM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                sFamilia = CType(oForm.Items.Item("cbFAM").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString
            Else
                sFamilia = ""
            End If
            If CType(oForm.Items.Item("cbSFAM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                sSFamilia = CType(oForm.Items.Item("cbSFAM").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString
            Else
                sSFamilia = ""
            End If
            sMarca = oForm.DataSources.UserDataSources.Item("UDMARCA").Value.ToString
            sAlmacen = oForm.DataSources.UserDataSources.Item("UDALM").Value.ToString
            sIC = oForm.DataSources.UserDataSources.Item("UDIC").Value.ToString
            oForm.Freeze(True)
            Select Case pVal.ItemUID
                Case "btnCa"
#Region "cargar Grid con los datos leidos"
                    'Ahora cargamos el Grid con los datos guardados
                    objGlobal.SBOApp.StatusBar.SetText("Cargando datos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    sSQL = " SELECT * FROM ("
                    sSQL &= "SELECT OITW.""WhsCode"" ""Almacén"",OWHS.""WhsName"" ""Desc. Almacén"", OITW.""ItemCode"" ""Artículo"", OITM.""ItemName"" ""Desc. Artículo"" "
                    sSQL &= " , OITB.""ItmsGrpNam"" ""Familia"", CONCAT(ifnull(""SUBFAM"".""U_EXO_FABDES"",''),CONCAT('-',ifnull(""SUBFAM"".""U_EXO_DESSUBFAM"",''))) ""Subfamilia"", OITM.""U_stec_marcas"" ""Marca"" "
                    sSQL &= ", OITW.""OnHand"" ""Stock"", OITW.""IsCommited"" ""comprometido"", OITW.""OnOrder"" ""Pedido"",ifnull(OMRC.""FirmName"",'') ""Fabricante"", Ifnull(DTO.""Discount"",0) ""Dto. IC"" "
                    sSQL &= " From OITW "
                    sSQL &= " INNER JOIN OWHS ON OITW.""WhsCode""=OWHS.""WhsCode"" "
                    sSQL &= " INNER JOIN OITM ON OITW.""ItemCode""=OITM.""ItemCode"" "
                    sSQL &= " LEFT JOIN OITB ON OITB.""ItmsGrpCod""=OITM.""ItmsGrpCod"" "
                    sSQL &= " LEFT JOIN OMRC ON OMRC.""FirmCode""=OITM.""FirmCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_FAMSUBFAM"" ""SUBFAM"" ON ""SUBFAM"".""DocEntry""=OITM.""U_EXO_SUBFAM"" "
                    sSQL &= " LEFT JOIN (SELECT EDG1.*, OMRC.""FirmName"" ""Fabricante"" FROM OEDG "
                    sSQL &= " LEFT JOIN EDG1 ON EDG1.""AbsEntry""=OEDG.""AbsEntry"" "
                    sSQL &= " LEFT JOIN OMRC ON OMRC.""FirmCode""=EDG1.""ObjKey"" "
                    sSQL &= " WHERE OEDG.""ObjCode""='" & sIC & "' and EDG1.""ObjType""='43' and ""ValidFor""='Y' "
                    sSQL &= " And (ifnull(""ValidForm"",'" & sFecha & "')>='" & sFecha & "' and ifnull(""ValidTo"",'" & sFecha & "')<='" & sFecha & "')) DTO ON DTO.""ObjKey""=OITM.""FirmCode"" "
                    sSQL &= ") t WHERE 1=1 "
                    If sCodArt.Trim <> "" Then
                        sSQL &= " and t.""Artículo""='" & sCodArt & "' "
                    End If
                    If sDesArt.Trim <> "" Then
                        sSQL &= " and t.""Desc. Artículo"" like '%" & sDesArt & "%' "
                    End If
                    If sFamilia.Trim <> "" And sFamilia.Trim <> "-" Then
                        sSQL &= " and t.""Familia"" ='" & sFamilia & "' "
                    End If
                    If sSFamilia.Trim <> "" And sSFamilia.Trim <> "-" Then
                        sSQL &= " and t.""Subfamilia"" ='" & sSFamilia & "' "
                    End If
                    If sMarca.Trim <> "" Then
                        sSQL &= " and t.""Marca"" like '%" & sMarca & "%' "
                    End If
                    If sAlmacen.Trim <> "" And sAlmacen.Trim <> "-" Then
                        sSQL &= " and t.""Almacén"" ='" & sAlmacen & "' "
                    End If

                    sSQL &= " ORDER BY t.""Artículo"", t.""Almacén"" "
                    'Cargamos grid
                    oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
                    FormateaGrid(oForm)
                    objGlobal.SBOApp.StatusBar.SetText("Datos cargados correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
            End Select
            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oform.Freeze(True)
            For i = 0 To 11
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                oColumnTxt.TitleObject.Sortable = True
                If (i >= 7 And i <= 9) Or i = 11 Then
                    oColumnTxt.RightJustified = True
                End If
            Next
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Function EventHandler_DOUBLE_CLICK_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim oFormDetalle As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim sFecha As String = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
        Dim sCodArt As String = ""

        Dim sIC As String = ""
        EventHandler_DOUBLE_CLICK_After = False

        Try
            If pVal.Row >= 0 Then
                sCodArt = oForm.DataSources.DataTables.Item("DT_DOC").GetValue("Artículo", pVal.Row).ToString
                sIC = oForm.DataSources.UserDataSources.Item("UDIC").Value.ToString

#Region "cargar Grid con los datos leidos"
                'Ahora cargamos el Grid con los datos guardados
                objGlobal.SBOApp.StatusBar.SetText("Cargando datos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = " SELECT * FROM ("
                sSQL &= "SELECT OITW.""WhsCode"" ""Almacén"",OWHS.""WhsName"" ""Desc. Almacén"", OITW.""ItemCode"" ""Artículo"", OITM.""ItemName"" ""Desc. Artículo"" "
                sSQL &= " , OITB.""ItmsGrpNam"" ""Familia"", ifnull(""SUBFAM"".""U_EXO_DESSUBFAM"",'') ""Subfamilia"", OITM.""U_stec_marcas"" ""Marca"" "
                sSQL &= ", OITW.""OnHand"" ""Stock"", OITW.""IsCommited"" ""comprometido"", OITW.""OnOrder"" ""Pedido"",ifnull(OMRC.""FirmName"",'') ""Fabricante"", Ifnull(DTO.""Discount"",0) ""Dto. IC"" "
                sSQL &= " From OITW "
                sSQL &= " INNER JOIN OWHS ON OITW.""WhsCode""=OWHS.""WhsCode"" "
                sSQL &= " INNER JOIN OITM ON OITW.""ItemCode""=OITM.""ItemCode"" "
                sSQL &= " LEFT JOIN OITB ON OITB.""ItmsGrpCod""=OITM.""ItmsGrpCod"" "
                sSQL &= " LEFT JOIN OMRC ON OMRC.""FirmCode""=OITM.""FirmCode"" "
                sSQL &= " LEFT JOIN ""@EXO_FAMSUBFAM"" ""SUBFAM"" ON ""SUBFAM"".""U_EXO_CODFAM""=OITM.""ItmsGrpCod"" and ""SUBFAM"".""U_EXO_CODFAM""=OITM.""U_EXO_SUBFAM"" "
                sSQL &= " LEFT JOIN (SELECT EDG1.*, OMRC.""FirmName"" ""Fabricante"" FROM OEDG "
                sSQL &= " LEFT JOIN EDG1 ON EDG1.""AbsEntry""=OEDG.""AbsEntry"" "
                sSQL &= " LEFT JOIN OMRC ON OMRC.""FirmCode""=EDG1.""ObjKey"" "
                sSQL &= " WHERE OEDG.""ObjCode""='" & sIC & "' and EDG1.""ObjType""='43' and ""ValidFor""='Y' "
                sSQL &= " And (ifnull(""ValidForm"",'" & sFecha & "')>='" & sFecha & "' and ifnull(""ValidTo"",'" & sFecha & "')<='" & sFecha & "')) DTO ON DTO.""ObjKey""=OITM.""FirmCode"" "
                sSQL &= ") t WHERE 1=1 "
                If sCodArt.Trim <> "" Then
                    sSQL &= " and t.""Artículo""='" & sCodArt & "' "
                End If
                sSQL &= " ORDER BY t.""Artículo"", t.""Almacén"" "
                Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
                oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
                oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CEART2.srf")
                oFP.XmlData = oFP.XmlData.Replace("modality=""0""", "modality=""1""")
                Try
                    oFormDetalle = objGlobal.SBOApp.Forms.AddEx(oFP)

                Catch ex As Exception
                    If ex.Message.StartsWith("Form - already exists") = True Then
                        objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                        Exit Function
                    End If
                End Try

                'Cargamos grid
                oFormDetalle.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
                FormateaGrid(oFormDetalle)
                objGlobal.SBOApp.StatusBar.SetText("Datos cargados correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
            End If



            EventHandler_DOUBLE_CLICK_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If oFormDetalle Is Nothing Then
            Else
                oFormDetalle.Visible = True
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oFormDetalle, Object))
        End Try
    End Function
End Class
