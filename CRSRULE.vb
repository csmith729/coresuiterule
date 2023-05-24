Imports System
Imports SwissAddonFramework.UI.Components
Imports SwissAddonFramework.UI.Dialogs
Imports SwissAddonFramework.Messaging
Imports customize = SwissAddonFramework
Imports Microsoft.VisualBasic
Imports ViewImgFromSapLIB
Imports System.Threading.Tasks
Imports Microsoft.Threading.Tasks
Imports Microsoft.Threading.Tasks.Extensions.Desktop
Imports Microsoft.Threading.Tasks.Extensions
Imports System.IO
Imports System.Runtime
'#if File.Exists(

Namespace COR_Utility
Class Helper

' Globals ----------
Function IsValidFileNameOrPath(ByVal name As String) As Boolean
	' Determines if the name is Nothing.
	If name Is Nothing Then
		Return False
	End If

	' Determines if there are bad characters in the name.
	For Each badChar As Char In System.IO.Path.GetInvalidPathChars
		If InStr(name, badChar) > 0 Then
			Return False
		End If
	Next

	' The name passes basic validation.
	Return True
End Function


Private Function CheckDLLExists(filePath As String) As Boolean
	' Check if the DLL file exists at the specified path
	Return System.IO.File.Exists(filePath)
End Function
' Function: Run_FO_COR_CUS_00000090 ----------
Public Function Run_FO_COR_CUS_00000090(pVal As COR_Utility.Logic.CustomizeEvent) As Boolean

'*/
Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application



Try
	StatusBar.WriteWarning("DEBUG - Rule: " + pVal.RuleInfo.RuleName + " was triggered.")
	' Your Code

	
	If application.Menus.Item("1280").SubMenus.Exists("1293") Then
						
		If application.Menus.Item("1280").SubMenus.Item("1293").Enabled = False Then application.Menus.Item("1280").SubMenus.Item("1293").Enabled = True
		
		
	End If
	
	If application.Menus.Item("1280").SubMenus.Exists("1293") Then
		
		If application.Menus.Item("1280").SubMenus.Item("1294").Enabled = False Then application.Menus.Item("1280").SubMenus.Item("1294").Enabled = True
		
		
	End If

Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	'MessageBox.Show(errorMessage, "OK")
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try

Return True
Return SwissAddonFramework.Global.SAPAction
End Function
' Function: Run_FO_COR_CUS_00000092 ----------
Public Function Run_FO_COR_CUS_00000092(pVal As COR_Utility.Logic.CustomizeEvent) As Boolean
'/*

Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)

Try
	StatusBar.WriteWarning("DEBUG - Rule: " + pVal.RuleInfo.RuleName + " was triggered.")
	StatusBar.WriteWarning("Form Unique ID: " + frm.UniqueID)
	
	
	' Your Code
Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	MessageBox.Show(errorMessage, "OK")
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try

Return True
Return SwissAddonFramework.Global.SAPAction
End Function

Public Function Run_FO_COR_CUS_00000091(pVal As COR_Utility.Logic.CustomizeEvent) As Boolean
Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim company As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)
Dim oItem As SAPbouiCOM.Item = frm.Items.Item("37")
Dim mtx As SAPbouiCOM.Matrix = oItem.Specific
Dim Col As SAPbouiCOM.Column
Dim ClkCol As SAPbouiCOM.Column
Dim Cll As SAPbouiCOM.Cell
Dim CellPos As SAPbouiCOM.CellPosition = mtx.GetCellFocus
Dim WH7TrfBalStr As String = "0"
Dim oEdit As SAPbouiCOM.EditText
Dim DecimalSep As String = System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator
Dim SelectedRowNumber(mtx.RowCount) As Integer
Dim k As Integer = 0
Dim TrfRowSum As Integer = 0
	
	
Try
	
	Col = mtx.Columns.Item(131)
	Statusbar.WriteWarning("pVal Row Index : " + pVal.Row.ToString)
	Statusbar.WriteWarning("pVal Object String : " + pVal.EventObject.ToString)
			
	If oItem.Enabled And mtx.RowCount <> 0 Then
		
		For p As Integer = 1 To mtx.RowCount - 1 Step 1
	
			If mtx.IsRowSelected(p) Then 
			
				WH7TrfBalStr = Col.Cells.Item(p).Specific.Value.ToString
				K += 1
			
				If Not WH7TrfBalStr.Contains(DecimalSep) And DecimalSep <> "." Then

					WH7TrfBalStr = WH7TrfBalStr.Replace(".", ",")
				

				End If
			
				TrfRowSum += Double.Parse(WH7TrfBalStr)
			
			End If
		
	
		Next
		
		

		If TrfRowSum = 0 Then 
		
				
			If pVal.Row <> mtx.RowCount Then
		
			
				If pVal.Row <> 0 And pVal.Row <> -1 Then 
				
					WH7TrfBalStr = Col.Cells.Item(pVal.Row).Specific.Value.ToString
		
					If Not WH7TrfBalStr.Contains(DecimalSep) And DecimalSep <> "." Then

						WH7TrfBalStr = WH7TrfBalStr.Replace(".", ",")
		

					End If
		
				End If
		
			End If

		End If
	
		Statusbar.WriteWarning("TrfRowSum : " + TrfRowSum.ToString)
		Statusbar.WriteWarning("WH7TrfBalStr : " + Double.Parse(WH7TrfBalStr).ToString)
	
		If Double.Parse(WH7TrfBalStr) > 0 Or TrfRowSum > 0 And pVal.Row <> 0 Then
		
			
					
			
			If application.Menus.Item("1280").SubMenus.Exists("1293") Then
					
				If application.Menus.Item("1280").SubMenus.Item("1293").Enabled = True Then application.Menus.Item("1280").SubMenus.Item("1293").Enabled = False
					
			End If
				
			If application.Menus.Item("1280").SubMenus.Exists("1294") Then
					
				If application.Menus.Item("1280").SubMenus.Item("1294").Enabled = True Then application.Menus.Item("1280").SubMenus.Item("1294").Enabled = False
					
			End If
				
		ElseIf Double.Parse(WH7TrfBalStr) = 0 And k = 0 And pVal.Row = -1 Then
			
			
			
			Return True
	

		
		Else
		
		

			If application.Menus.Item("1280").SubMenus.Exists("1293") Then
				
				If application.Menus.Item("1280").SubMenus.Item("1293").Enabled = False Then application.Menus.Item("1280").SubMenus.Item("1293").Enabled = True
				
			End If
			
			If application.Menus.Item("1280").SubMenus.Exists("1294") Then
				
				If application.Menus.Item("1280").SubMenus.Item("1294").Enabled = False Then application.Menus.Item("1280").SubMenus.Item("1294").Enabled = True
				
			End If
			

		End If
		
		
	End If

Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	'MessageBox.Show(errorMessage, "OK")
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try

Return True
Return SwissAddonFramework.Global.SAPAction
End Function
' Function: Run_FO_COR_CUS_00000086 ----------
Public Function Run_FO_COR_CUS_00000086(pVal As COR_Utility.Logic.CustomizeEvent) As Boolean
Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim company As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)
Dim oItem As SAPbouiCOM.Item = frm.Items.Item("37")
Dim pItem As SAPbouiCOM.Form
Dim PrdNumStr As SAPbouiCOM.Item = frm.Items.Item(18)
Dim mtx As SAPbouiCOM.Matrix = oItem.Specific
Dim Col As SAPbouiCOM.Column
Dim TxBxPrdNum As SAPbouiCOM.EditText
Dim IssuedCol As SAPbouiCOM.Column
Dim ItmCol As SAPbouiCOM.Column
Dim ClkCol As SAPbouiCOM.Column = mtx.Columns.Item(pVal.ColUID)
Dim Cll As SAPbouiCOM.Cell
Dim CellPos As SAPbouiCOM.CellPosition = mtx.GetCellFocus
Dim WH7TrfBalStr As String = "0"
Dim oEdit As SAPbouiCOM.EditText
Dim DecimalSep As String = System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator
Dim SelectedRowNumber(mtx.RowCount) As Integer
Dim k As Integer = 0
Dim TrfRowSum As Integer = 0
Dim IssuedQtySum As Integer = 0
	
	
Try
	
	TxBxPrdNum = frm.Items.Item("18").Specific
	Col = mtx.Columns.Item(131)
	ItmCol = mtx.Columns.Item("4")
	IssuedCol = mtx.Columns.Item("13")
	Statusbar.WriteWarning("pVal Row Index : " + pVal.Row.ToString)
	Statusbar.WriteWarning("pVal Object String : " + pVal.EventObject.ToString)
			
	
	
	If oItem.Enabled And mtx.RowCount <> 0 Then
		
		For p As Integer = 1 To mtx.RowCount - 1 Step 1
	
			If mtx.IsRowSelected(p) Then 
			
				WH7TrfBalStr = Col.Cells.Item(p).Specific.Value.ToString
				
			
				If Not WH7TrfBalStr.Contains(DecimalSep) And DecimalSep <> "." Then

					WH7TrfBalStr = WH7TrfBalStr.Replace(".", ",")
				

				End If
			
				TrfRowSum += Double.Parse(WH7TrfBalStr)
				IssuedQtySum += IssuedCol.Cells.Item(p).Specific.Value
			
			End If
		
	
		Next

	
		If TrfRowSum = 0 Then 
		
				
			If pVal.Row <> mtx.RowCount Then
		
			
				If pVal.Row <> 0 Then 
				
					WH7TrfBalStr = Col.Cells.Item(pVal.Row).Specific.Value.ToString
					
		
					If Not WH7TrfBalStr.Contains(DecimalSep) And DecimalSep <> "." Then

						WH7TrfBalStr = WH7TrfBalStr.Replace(".", ",")

					End If
		
				End If
		
			End If
	
	
		End If
	
		Statusbar.WriteWarning("TrfRowSum : " + TrfRowSum.ToString)
		Statusbar.WriteWarning("WH7TrfBalStr : " + Double.Parse(WH7TrfBalStr).ToString)
		Statusbar.WriteWarning("Issued QTYSUM : " + IssuedQtySum.ToString + " Prd Num: " + TxBxPrdNum.String) ' 
	
		If Double.Parse(WH7TrfBalStr) > 0 Or TrfRowSum > 0 And pVal.Row <> 0 And Not ClkCol.Editable Then
		
					
			
			If application.Menus.Item("1280").SubMenus.Exists("1293") Then
					
				If application.Menus.Item("1280").SubMenus.Item("1293").Enabled = True Then application.Menus.Item("1280").SubMenus.Item("1293").Enabled = False
					
			End If
				
			If application.Menus.Item("1280").SubMenus.Exists("1294") Then
					
				If application.Menus.Item("1280").SubMenus.Item("1294").Enabled = True Then application.Menus.Item("1280").SubMenus.Item("1294").Enabled = False
					
			End If
				
				
			If application.Menus.Exists("774") Then
					
				If application.Menus.Item("774").Enabled = True Then application.Menus.Item("774").Enabled = False
					
			End If
				
			If application.Menus.Exists("773") Then
					
				If application.Menus.Item("773").Enabled = True Then application.Menus.Item("773").Enabled = False
					
			End If
			

	

		
		Else
		
		

			If application.Menus.Item("1280").SubMenus.Exists("1293") Then
				
				If application.Menus.Item("1280").SubMenus.Item("1293").Enabled = False Then application.Menus.Item("1280").SubMenus.Item("1293").Enabled = True
				
			End If
			
			If application.Menus.Item("1280").SubMenus.Exists("1294") Then
				
				If application.Menus.Item("1280").SubMenus.Item("1294").Enabled = False Then application.Menus.Item("1280").SubMenus.Item("1294").Enabled = True
				
			End If
			
			
			If application.Menus.Exists("774") Then
				
				If application.Menus.Item("774").Enabled = False Then application.Menus.Item("774").Enabled = True
				
			End If
			
			If application.Menus.Exists("773") Then
				
				If application.Menus.Item("773").Enabled = False Then application.Menus.Item("773").Enabled = True
				
		
			
			End If

		

		
		
		End If
		
		
	End If
		

	
	
	
Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try

Return True
Return SwissAddonFramework.Global.SAPAction
End Function
' Function: Run_FO_COR_CUS_00000098 ----------
Public Function Run_FO_COR_CUS_00000098(pVal As COR_Utility.Logic.CustomizeEvent) As Boolean
'/*


    If File.Exists("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") = True Then

        Imports Microsoft.Office.Interop
        Imports Microsoft.Office.Interop.Excel

        Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
        Dim company As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
        Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)
        Dim LineItemTab As SAPbouiCOM.Folder = frm.Items.Item("112").Specific
        Dim LogisticsTab As SAPbouiCOM.Folder = frm.Items.Item("114").Specific
        Dim AccountingTab As SAPbouiCOM.Folder = frm.Items.Item("138").Specific
        Dim ElecDocsTab As SAPbouiCOM.Folder = frm.Items.Item("350002087").Specific
        Dim AttachmentPane As SAPbouiCOM.Folder = frm.Items.Item("BOYT_1").Specific
        Dim ReturnTab As SAPbouiCOM.Folder
        Dim oItem As SAPbouiCOM.Item = frm.Items.Item("38")
        Dim mtx As SAPbouiCOM.Matrix = oItem.Specific
        Dim Remks As SAPbouiCOM.EditText = frm.Items.Item("16").Specific
        Dim CmbxSalesType As SAPbouiCOM.ComboBox = frm.Items.Item("NI_00024").Specific
        Dim ET_QT_Cost_Path1 As SAPbouiCOM.EditText = frm.Items.Item("BOYX_1").Specific
        Dim MustReturn As Boolean = False
        Dim ItemsMaster As SAPbobsCOM.Items = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

        Try

            Dim QuotedCostPriceCol As SAPbouiCOM.Column = mtx.Columns.Item("U_QuotedPrice")
            Dim ItemCodeCol As SAPbouiCOM.Column = mtx.Columns.Item("1")
            Dim QuotedCostMatCostCol As SAPbouiCOM.Column = mtx.Columns.Item("U_CostingCostPrice")
            Dim QuotedCostPrdNameCol As SAPbouiCOM.Column = mtx.Columns.Item("U_CostingPrdName")
            Dim QuotedCostQTYCol As SAPbouiCOM.Column = mtx.Columns.Item("U_CostingQTY")
            Dim QuotedCostBOMCodeCol As SAPbouiCOM.Column = mtx.Columns.Item("U_CostingBOMCode")
            Dim QuotedCostLabCol As SAPbouiCOM.Column = mtx.Columns.Item("U_CostingLabourCost")
            Dim QuotedLineCostAttachCol As SAPbouiCOM.Column = mtx.Columns.Item("U_LineCostAttachment")
            Dim RowHdrCol As SAPbouiCOM.Column = mtx.Columns.Item("0")
            Dim FialedLineNumber As Integer = -1
            Dim NoCostAttached As Boolean = True
            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheetRM As Excel.Worksheet
            Dim xlWorkSheetCS As Excel.Worksheet

            For i As Integer = 1 To mtx.RowCount - 1 Step 1

                NoCostAttached = True

                If ItemsMaster.GetByKey(ItemCodeCol.Cells.Item(i).Specific.Value.ToString) Then

                    If ItemsMaster.TreeType = 3 Then

                        If Len(QuotedLineCostAttachCol.Cells.Item(i).Specific.Value.ToString) < 5 Then
                            StatusBar.WriteWarning("DEBUG - Rule: " & pVal.RuleInfo.RuleName & " was triggered. -" & QuotedLineCostAttachCol.Cells.Item(i).Specific.Value.ToString)
                            MessageBox.Show("You must add a cost attachment", "OK")
                            FialedLineNumber = i
                            NoCostAttached = True
                            Exit For
                        End If

                        StatusBar.WriteWarning("Path of Line Nuber" & i & ": " & "Path : " & QuotedLineCostAttachCol.Cells.Item(i).Specific.Value.ToString)

                        Dim FileInf As New FileInfo(QuotedLineCostAttachCol.Cells.Item(i).Specific.Value.ToString)

                        If FileInf.Exists Then

                            If FileInf.Extension <> ".xlsm" Then

                                MessageBox.Show("The Cost Attachment must be of file type xlms", "OK")
                                FialedLineNumber = i
                                NoCostAttached = True
                                Exit For

                            End If

                            StatusBar.WriteWarning("DEBUG - Rule: " & pVal.RuleInfo.RuleName & " was triggered.")

                            frm.Freeze(True)
                            NoCostAttached = False

                            If LineItemTab.Selected = False Then
                                If LogisticsTab.Selected = True Then
                                    ReturnTab = LogisticsTab
                                ElseIf AccountingTab.Selected = True Then
                                    ReturnTab = AccountingTab
                                ElseIf ElecDocsTab.Selected = True Then
                                    ReturnTab = ElecDocsTab
                                ElseIf AttachmentPane.Selected = True Then
                                    ReturnTab = AttachmentPane
                                End If
                                MustReturn = True
                                LineItemTab.Select()
                            End If

                            xlWorkBook = xlApp.Workbooks.Open(QuotedLineCostAttachCol.Cells.Item(i).Specific.Value.ToString)

                            Dim FoundRMSh As Boolean = False
                            Dim FoundCSSh As Boolean = False

                            For Each ws As Excel.Worksheet In xlWorkBook.Sheets

                                If ws.Name = "RM MOQ ADDED" Then

                                    xlWorkSheetRM = ws
                                    FoundRMSh = True

                                End If

                                If ws.Name = "Product Cost Summary" Then

                                    xlWorkSheetCS = ws
                                    FoundCSSh = True

                                End If

                            Next

                            If Not FoundRMSh And Not FoundCSSh Then

                                frm.Freeze(False)
                                MessageBox.Show("The Cost Attachment is not in the correct format", "OK")
                                FialedLineNumber = i
                                NoCostAttached = True
                                Exit For

                            End If

                            If xlWorkSheetRM.Range("E4").Value.ToString <> ItemCodeCol.Cells.Item(i).Specific.Value.ToString Then

                                frm.Freeze(False)
                                MessageBox.Show("The BOM number of the Attached costing does not match", "OK")
                                FialedLineNumber = i
                                NoCostAttached = True
                                Exit For

                            End If

                            StatusBar.WriteWarning("DEBUG - Rule: " & pVal.RuleInfo.RuleName & " " & QuotedLineCostAttachCol.Cells.Item(i).Specific.Value.ToString & " was opened.")

                            Remks.Active = True

                            If Not QuotedCostPriceCol.Visible Then QuotedCostPriceCol.Visible = True
                            If Not QuotedCostPriceCol.Editable Then QuotedCostPriceCol.Editable = True
                            If Not QuotedCostMatCostCol.Visible Then QuotedCostMatCostCol.Visible = True
                            If Not QuotedCostMatCostCol.Editable Then QuotedCostMatCostCol.Editable = True
                            If Not QuotedCostPrdNameCol.Visible Then QuotedCostPrdNameCol.Visible = True
                            If Not QuotedCostPrdNameCol.Editable Then QuotedCostPrdNameCol.Editable = True
                            If Not QuotedCostQTYCol.Visible Then QuotedCostQTYCol.Visible = True
                            If Not QuotedCostQTYCol.Editable Then QuotedCostQTYCol.Editable = True
                            If Not QuotedCostBOMCodeCol.Visible Then QuotedCostBOMCodeCol.Visible = True
                            If Not QuotedCostBOMCodeCol.Editable Then QuotedCostBOMCodeCol.Editable = True
                            If Not QuotedCostLabCol.Visible Then QuotedCostLabCol.Visible = True
                            If Not QuotedCostLabCol.Editable Then QuotedCostLabCol.Editable = True

                            QuotedCostPrdNameCol.Cells.Item(i).Specific.Value = xlWorkSheetRM.Range("E5").Value
                            QuotedCostMatCostCol.Cells.Item(i).Specific.Value = xlWorkSheetRM.Range("P8").Value
                            QuotedCostQTYCol.Cells.Item(i).Specific.Value = xlWorkSheetRM.Range("E6").Value
                            QuotedCostPriceCol.Cells.Item(i).Specific.Value = xlWorkSheetCS.Range("G39").Value
                            QuotedCostBOMCodeCol.Cells.Item(i).Specific.Value = xlWorkSheetRM.Range("E4").Value
                            QuotedCostLabCol.Cells.Item(i).Specific.Value = xlWorkSheetCS.Range("G12").Value

                            Remks.Active = True
                            QuotedCostPriceCol.Visible = False
                            QuotedCostPriceCol.Editable = False
                            QuotedCostMatCostCol.Visible = False
                            QuotedCostMatCostCol.Editable = False
                            QuotedCostPrdNameCol.Visible = False
                            QuotedCostPrdNameCol.Editable = False
                            QuotedCostQTYCol.Visible = False
                            QuotedCostQTYCol.Editable = False
                            QuotedCostBOMCodeCol.Visible = False
                            QuotedCostBOMCodeCol.Editable = False
                            QuotedCostLabCol.Visible = False
                            QuotedCostLabCol.Editable = False

                            If MustReturn Then

                                ReturnTab.Select()

                            End If

                            frm.Freeze(False)

                        End If

                    End If

                Next

                If xlApp IsNot Nothing Then

                    xlWorkBook.Close()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheetRM)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheetCS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook)
                    xlWorkSheetRM = Nothing
                    xlWorkSheetCS = Nothing
                    xlWorkBook = Nothing

                    xlApp.Quit()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
                    xlApp = Nothing

                End If

                If NoCostAttached = True Then
                    RowHdrCol.Cells.Item(FialedLineNumber).Click()
                    StatusBar.WriteError("No Cost attachment for Line: " & FialedLineNumber)
                    Return False
                End If
            Catch ex As Exception
                Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
                frm.Freeze(False)
                MessageBox.Show(errorMessage, "OK")
                StatusBar.WriteError(errorMessage)
                Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
            End Try

            Return True

        Else

            Return True
        End If
Return SwissAddonFramework.Global.SAPAction
End Function
' Function: Run_FO_COR_CUS_00000089 ----------
Public Function Run_FO_COR_CUS_00000089(pVal As COR_Utility.Logic.CustomizeEvent) As Boolean
'/*

Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim company As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)
Dim oRemkrs As SAPbouiCOM.Item = frm.Items.Item("11")
Dim oItem As SAPbouiCOM.Item = frm.Items.Item("13")
Dim mtx As SAPbouiCOM.Matrix = oItem.Specific
 
Try
	StatusBar.WriteWarning("DEBUG - Rule: " + pVal.RuleInfo.RuleName + " was triggered.")
	Dim Wh7Col As SAPbouiCOM.Column
	Dim IssuWhCol As SAPbouiCOM.Column
	
	frm.Freeze(True)
	
	For i As Integer = 0 To mtx.Columns.Count - 1 Step 1
		
		If mtx.Columns.Item(i).UniqueID = "U_MTX_TRANSFERRED" Then
			
			Wh7Col = mtx.Columns.Item(i)
			
		End If
		
		If mtx.Columns.Item(i).UniqueID = "15" Then
			
			IssuWhCol = mtx.Columns.Item(i)
			
			
		End If

	Next

	
	Dim DelRwCnt As Integer = 1
	
	
		
	For p As Integer = mtx.RowCount - 1 To 0 Step -1
		

		
		If Wh7Col.Cells.Item(p + 1).Specific.Value.ToString = "0.0" Or Wh7Col.Cells.Item(p + 1).Specific.Value.ToString = "0,0" Then
			
			mtx.DeleteRow(p + 1)

		Else
			
			IssuWhCol.Cells.Item(p + 1).Specific.Value = "07"

		End If
		
		
	Next
	
	Wh7Col.Editable = False
	oRemkrs.Click
	frm.Freeze(False)
	
	' Your Code
Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	MessageBox.Show(errorMessage, "OK")
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try

Return True
Return SwissAddonFramework.Global.SAPAction
End Function

Public Function Run_FO_COR_CUS_00000085(pVal As COR_Utility.Logic.CustomizeEvent) As Boolean


Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim company As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)

Dim oItem As SAPbouiCOM.Item = frm.Items.Item("37")
Dim oStatus As SAPbouiCOM.Item = frm.Items.Item("10")
Dim oRemkrs As SAPbouiCOM.Item = frm.Items.Item("3")
Dim oUpdateBtn As SAPbouiCOM.Item = frm.Items.Item("1")
Dim mtx As SAPbouiCOM.Matrix = oItem.Specific
Dim DecimalSep As String = System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator



Try
	
	Dim DocEntry As String = TextEdit.GetFromUID(pVal.Form, "18").Value.ToString
	Dim PrdOrdDueDate As SAPbouiCOM.Item = frm.Items.Item("26")
	Dim Col As SAPbouiCOM.Column = mtx.Columns.Item(3)
	Dim UCol As SAPbouiCOM.Column = mtx.Columns.Item(131)
	Dim PldQTYCol As SAPbouiCOM.Column = mtx.Columns.Item(8)
	Dim EndDateCol As SAPbouiCOM.Column = mtx.Columns.Item("2340000045")
	Dim TrfUpdated As Boolean = False
	Dim DateUpdated As Boolean = False
	Dim PldQTYNum As Double = 0
	Dim Wh7QtyNum As Double = 0

	
	For o As Integer = 0 To mtx.RowCount - 1 Step 1
			
		If EndDateCol.Cells.Item(o + 1).Specific.Value > PrdOrdDueDate.Specific.Value Then
					
			EndDateCol.Cells.Item(o + 1).Specific.Value = PrdOrdDueDate.Specific.Value
			DateUpdated = True
			StatusBar.WriteWarning("Form Row Number: " + o.ToString + " Date Changed")
					
		End If
			
			
	Next
	

	If oItem.Enabled Then

		mtx.Columns.Item("15").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
		
		Dim query1 As String = "Select T1.[ItemCode], SUM(Case When T1.[FromWhsCod] = '07' THEN T1.[Quantity] * -1 ELSE T1.[Quantity] * 1 END) AS [QTY] INTO ITRFBAL FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[U_PROD_NO] =  " + DocEntry + "  AND (T1.[FromWhsCod] = '07' OR T1.[WhsCode] = '07') GROUP BY T1.[ItemCode] SELECT T2.[ItemCode], T2.[QTY], T3.[LineNum]  FROM ITRFBAL T2 INNER JOIN WOR1 T3 ON T2.[ItemCode] = T3.[ItemCode] WHERE T3.[DocEntry] = " + DocEntry + " ORDER BY T3.[LineNum] ASC DROP TABLE ITRFBAL"
		Dim query2 As String = "SELECT T2.[LineNum], T2.[ItemCode] , SUM(Case When T2.[ItemCode] = T0.[ItemCode] And T0.[FromWhsCod] = '07' AND T1.[U_PROD_NO] = T2.[DocEntry] THEN -T0.[Quantity] WHEN T2.[ItemCode] = T0.[ItemCode] AND T0.[WhsCode] = '07' AND T1.[U_PROD_NO] = T2.[DocEntry] THEN T0.[Quantity] ELSE 0 END) AS [WH7 QTY] INTO #WH7TRFQTY FROM WTR1 T0 INNER JOIN OWTR T1 ON T1.[DocEntry] = T0.[DocEntry] RIGHT JOIN WOR1 T2 ON T2.[DocEntry] = T1.[U_PROD_NO] WHERE T2.[DocEntry] = " + DocEntry + " GROUP BY T2.[ItemCode], T2.[LineNum] ORDER BY T2.[LineNum] ASC SELECT T1.[ItemCode], SUM(COALESCE(T0.[Quantity], 0)) AS [Issued QTY] INTO #ISSUEDQTY FROM IGE1 T0 RIGHT JOIN WOR1 T1 ON CONCAT(T1.[DocEntry], T1.[LineNum]) = CONCAT(T0.[BaseEntry], T0.[BaseLine]) WHERE T1.[DocEntry] = " + DocEntry + " AND T1.[ItemCode] NOT LIKE 'MLAB%' GROUP BY T1.[ItemCode] SELECT T1.[ItemCode], SUM(COALESCE(T0.[Quantity], 0)) AS [Return QTY] INTO #RETURNQTY FROM IGN1 T0 RIGHT JOIN WOR1 T1 ON CONCAT(T1.[DocEntry], T1.[LineNum]) = CONCAT(T0.[BaseEntry], T0.[BaseLine]) WHERE T1.[DocEntry] = " + DocEntry + " AND T1.[ItemCode] NOT LIKE 'MLAB%' GROUP BY T1.[ItemCode] SELECT #WH7TRFQTY.[LineNum], #WH7TRFQTY.[ItemCode], #WH7TRFQTY.[WH7 QTY], #ISSUEDQTY.[Issued QTY], #RETURNQTY.[Return QTY] FROM #WH7TRFQTY INNER JOIN #ISSUEDQTY ON #ISSUEDQTY.[ItemCode] = #WH7TRFQTY.[ItemCode] INNER JOIN #RETURNQTY ON #RETURNQTY.[ItemCode] = #ISSUEDQTY.[ItemCode] DROP TABLE #WH7TRFQTY, #ISSUEDQTY, #RETURNQTY"
		Dim query3 As String = "WITH CTE_WH7TRFQTY AS (    SELECT T2.[LineNum], T2.[ItemCode],            SUM(CASE                   WHEN T2.[ItemCode] = T0.[ItemCode] AND T0.[FromWhsCod] = '07' AND T1.[U_PROD_NO] = T2.[DocEntry] THEN -T0.[Quantity]                   WHEN T2.[ItemCode] = T0.[ItemCode] AND T0.[WhsCode] = '07' AND T1.[U_PROD_NO] = T2.[DocEntry] THEN T0.[Quantity]                   ELSE 0               END) AS [WH7 QTY]    FROM WTR1 T0     INNER JOIN OWTR T1     ON T1.[DocEntry] = T0.[DocEntry]     RIGHT JOIN WOR1 T2     ON T2.[DocEntry] =  T1.[U_PROD_NO]    WHERE T2.[DocEntry] = " + DocEntry + " AND T1.[DocDate] >= '02/26/2023'    GROUP BY T2.[ItemCode], T2.[LineNum]),CTE_ISSUEDQTY AS (    SELECT T1.[ItemCode], SUM(ISNULL( T0.[Quantity], 0)) AS [Issued QTY]     FROM IGE1 T0     RIGHT JOIN WOR1 T1     ON CONCAT(T1.[DocEntry], T1.[LineNum])  = CONCAT(T0.[BaseEntry], T0.[BaseLine])     WHERE T1.[DocEntry] = " + DocEntry + " AND T1.[ItemCode] NOT LIKE 'MLAB%'     GROUP BY T1.[ItemCode]),CTE_RETURNQTY AS (    SELECT T1.[ItemCode], SUM(ISNULL( T0.[Quantity], 0)) AS [Return QTY]     FROM IGN1 T0     RIGHT JOIN WOR1 T1     ON CONCAT(T1.[DocEntry], T1.[LineNum])  = CONCAT(T0.[BaseEntry], T0.[BaseLine])     WHERE T1.[DocEntry] = " + DocEntry + " AND T1.[ItemCode] NOT LIKE 'MLAB%'     GROUP BY T1.[ItemCode])SELECT CTE_WH7TRFQTY.[LineNum], CTE_WH7TRFQTY.[ItemCode], CTE_WH7TRFQTY.[WH7 QTY], CTE_ISSUEDQTY.[Issued QTY], CTE_RETURNQTY.[Return QTY] FROM CTE_WH7TRFQTY INNER JOIN CTE_ISSUEDQTY ON CTE_ISSUEDQTY.[ItemCode] = CTE_WH7TRFQTY.[ItemCode]INNER JOIN CTE_RETURNQTY ON CTE_RETURNQTY.[ItemCode] = CTE_ISSUEDQTY.[ItemCode]ORDER BY CTE_WH7TRFQTY.[LineNum] ASC"

		rs.DoQuery(query3)
		UCol.Editable = True
		
		Dim Wh7Val As String = UCol.Cells.Item(1).Specific.Value.ToString
		Dim PldQTYStr As String = PldQTYCol.Cells.Item(1).Specific.Value.ToString

		If	rs.RecordCount <> 0 Then
		
			Dim LstItmCode As New System.Collections.Generic.List(Of String)
			Dim LstItmQty As New System.Collections.Generic.List(Of Double)
			Dim LstItemRetQty As New System.Collections.Generic.List(Of Double)
			Dim LstItemIssQty As New System.Collections.Generic.List(Of Double)
			Dim LstItmRwNbr As New System.Collections.Generic.List(Of Double)
			rs.MoveFirst
			
			For g As Integer = 0 To rs.RecordCount - 1 Step 1

				LstItmCode.Add(rs.Fields.Item(1).Value.ToString)
				LstItemRetQty.Add(rs.Fields.Item(4).Value)
				LstItemIssQty.Add(rs.Fields.Item(3).Value)
				LstItmQty.Add(rs.Fields.Item(2).Value)
				LstItmRwNbr.Add(rs.Fields.Item(0).Value)
				rs.MoveNext

			Next

			For i As Integer = 0 To mtx.RowCount - 1 Step 1
				For k As Integer = 0 To LstItmCode.Count - 1 Step 1
					If LstItmQty.Item(k) = LstItemRetQty.Item(k) And LstItmQty.Item(k) = LstItemIssQty.Item(k) Then
						
						LstItmQty.Item(k) = 0
						
					End If
					If Left(Col.Cells.Item(i + 1).Specific.Value.ToString, 4) = "MLAB" Or LstItmRwNbr.Contains(mtx.Columns.Item("15").Cells.Item(i + 1).Specific.Value - 1) = False Then
						
						Exit For
						
					End If
					
					If LstItmCode.Item(k) = Col.Cells.Item(i + 1).Specific.Value Then 
						
						Wh7Val = UCol.Cells.Item(i + 1).Specific.Value.ToString
						
						If Not Wh7Val.Contains(DecimalSep) And DecimalSep <> "." Then
							Wh7Val = Wh7Val.Replace(".", ",")
						End If
						If Double.Parse(Wh7Val) <> LstItmQty.Item(k) Then
						
							
							UCol.Cells.Item(i + 1).Specific.Value = LstItmQty.Item(k)
							TrfUpdated = True
						
						End If
							
						LstItmCode.RemoveAt(k)
						LstItmQty.RemoveAt(k)
						LstItemRetQty.RemoveAt(k)
						LstItemIssQty.RemoveAt(k)
						LstItmRwNbr.RemoveAt(k)
						
						Exit For
						
					End If
				Next
				Wh7Val = UCol.Cells.Item(i + 1).Specific.Value.ToString
				PldQTYStr = PldQTYCol.Cells.Item(i + 1).Specific.Value.ToString

				If Not PldQTYStr.Contains(DecimalSep) And DecimalSep <> "." Then

					Wh7Val = Wh7Val.Replace(".", ",")
					PldQTYStr = PldQTYStr.Replace(".", ",")

				End If
				
				PldQTYNum = Double.Parse(PldQTYStr)
				Wh7QtyNum = Double.Parse(Wh7Val)
				
				If PldQTYNum <= Wh7QtyNum And Left(Col.Cells.Item(i + 1).Specific.Value.ToString, 4) <> "MLAB" And Not String.IsNullOrEmpty(Col.Cells.Item(i + 1).Specific.Value.ToString) Then
					mtx.CommonSetting.SetCellBackColor(i + 1, 8, rgb(116, 232, 139)) 'Green
				ElseIf Wh7QtyNum = 0 And Left(Col.Cells.Item(i + 1).Specific.Value.ToString, 4) <> "MLAB" And Not String.IsNullOrEmpty(Col.Cells.Item(i + 1).Specific.Value.ToString) Then
					mtx.CommonSetting.SetCellBackColor(i + 1, 8, rgb(236, 116, 116)) 'Red
				ElseIf Left(Col.Cells.Item(i + 1).Specific.Value.ToString, 4) <> "MLAB" And Not String.IsNullOrEmpty(Col.Cells.Item(i + 1).Specific.Value.ToString)
					mtx.CommonSetting.SetCellBackColor(i + 1, 8, rgb(255, 179, 0)) 'Orange
				Else
					mtx.CommonSetting.SetCellBackColor(i + 1, 8, rgb(255, 255, 255)) 'White
				End If

			
			Next

			oRemkrs.Click

			If TrfUpdated Or DateUpdated Then 
				oUpdateBtn.Click
			End If
		
			UCol.Editable = False
		
		Else
			
			UCol.Editable = False
			For i As Integer = 0 To mtx.RowCount - 1 Step 1
				
				mtx.CommonSetting.SetCellBackColor(i + 1, 8, rgb(255, 255, 255)) 'White

			Next

		End If
	
	Else
		
		UCol.Editable = False
		
		For i As Integer = 0 To mtx.RowCount - 1 Step 1
				
			mtx.CommonSetting.SetCellBackColor(i + 1, 8, rgb(230, 230, 230)) 'Gray
				
				
		Next

			
	End If

Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try


Return True
Return SwissAddonFramework.Global.SAPAction
End Function

Public Sub Run_FB_COR_CUS_00000041(pVal As COR_Utility.Logic.CustomizeEvent)

Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim oCompany As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
Dim ps As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
Dim RunningTotal As Integer = 0
Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)
Dim SalesOrderDocNum As SAPbouiCOM.EditText = frm.Items.Item("8").Specific
Dim CustomerNumber As SAPbouiCOM.EditText = frm.Items.Item("4").Specific
Dim EstDelDatePkr As SAPbouiCOM.EditText = frm.Items.Item("12").Specific
Dim EstRelDatePkr As SAPbouiCOM.EditText = frm.Items.Item("46").Specific
Dim oItem As SAPbouiCOM.Item = frm.Items.Item("38")
Dim mtx As SAPbouiCOM.Matrix = oItem.Specific
Dim ItemColumn As SAPbouiCOM.Column = mtx.Columns.Item("1")
Dim QTYColumn As SAPbouiCOM.Column = mtx.Columns.Item("11")
Dim ProdOrdColumn As SAPbouiCOM.Column = mtx.Columns.Item("U_QYC_PONR")
Dim ProjectColumn As SAPbouiCOM.Column = mtx.Columns.Item("31")
Dim CostCentreColumn As SAPbouiCOM.Column = mtx.Columns.Item("110000310")
Dim DelDateColumn As SAPbouiCOM.Column = mtx.Columns.Item("25")
Dim oProductionOrd As SAPbobsCOM.ProductionOrders
Dim oMBOM As SAPbobsCOM.ProductTrees
Dim oSalesOrder As SAPbobsCOM.Documents
Dim oItmMast As SAPbobsCOM.Items
Dim oSalesOrderLines As SAPbobsCOM.Document_Lines


Try
	
	If frm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
		
		
		StatusBar.WriteWarning("The Sales Order must be in 'OK Mode'")
		Exit Sub
		
	
	End If
	
	
	If Len(ProdOrdColumn.Cells.Item(pVal.Row).Specific.Value) > 0 Then
		
		StatusBar.WriteWarning("There is already a production order linked to this line")
		Exit Sub
		
		
	End If
	
	

	If pVal.Row <> -1 Then
		
		Dim BOMCode As String = ItemColumn.Cells.Item(pVal.Row).Specific.Value
		Dim SalesOrderNumber As Integer = Integer.Parse(SalesOrderDocNum.Value)
		Dim CustomerAccNbr As String = CustomerNumber.Value
		Dim CostCentreStr As String = CostCentreColumn.Cells.Item(pVal.Row).Specific.Value
		Dim QTYStr As String = QTYColumn.Cells.Item(pVal.Row).Specific.Value
		Dim PrjCodeStr As String = ProjectColumn.Cells.Item(pVal.Row).Specific.Value
		Dim SalesORderQTY As Integer = -1
		Dim InternalDocNum As Integer = -1
		
		

		
		
		If QTYStr.Contains(",") Then
			
			QTYStr.Replace(",", "")
			
		End If
		
		Dim SubBOMQuery As String = "DECLARE @FatherBOM NVARCHAR(20) = '" + BOMCode + "';  IF OBJECT_ID('tempdb..#BOMTable') IS NOT NULL     DROP TABLE #BOMTable;  WITH BOM_CTE AS (      SELECT          L1.Code AS Father,         L2.ItemName AS Father_Name,         L1.Code AS Child,         L2.ItemName AS Child_Name,         L2.TreeType AS BOMType,         1 AS Level,         CAST(L1.Quantity AS FLOAT) * " + QTYStr + " AS Quantity,         L2.InvntryUom AS UOM     FROM ITT1 L1     JOIN OITM L2 ON L1.Code = L2.ItemCode     WHERE L1.Father = @FatherBOM     UNION ALL      SELECT          L1.Code AS Father,         L2.ItemName AS Father_Name,         L1.Code AS Child,         L2.ItemName AS Child_Name,         L2.TreeType AS BOMType,         Level + 1 AS Level,         CAST(L1.Quantity AS FLOAT) * BOM_CTE.Quantity AS Quantity,         L2.InvntryUom AS UOM     FROM ITT1 L1     JOIN OITM L2 ON L1.Code = L2.ItemCode     JOIN BOM_CTE ON L1.Father = BOM_CTE.Child     WHERE L2.TreeType = 'P' )  SELECT      Father,     Father_Name,     Child,     Child_Name,     BOMType,     Level,     Quantity,     UOM INTO #BOMTable FROM BOM_CTE OPTION (MAXRECURSION 32767);     SELECT #BOMTable.Father, SUM(#BOMTable.Quantity) AS [Req QTY] FROM #BOMTable WHERE #BOMTable.BOMType = 'P' GROUP BY #BOMTable.Father;"
		
		'SalesORderQTY = Integer.Parse(QTYStr)
		
		oProductionOrd = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
		
		oSalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
		oMBOM = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
		oItmMast = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
		
		rs.DoQuery("SELECT TOP 1 T0.[DocEntry] FROM ORDR T0 WHERE T0.[DocNum] = " + SalesOrderNumber.ToString)
		
		
		
		If rs.RecordCount = 1 Then
			
			rs.MoveFirst
			
			InternalDocNum = rs.Fields.Item(0).Value
			
		Else
			
			Exit Sub

		End If

		If oSalesOrder.GetByKey(InternalDocNum) = False Then
			
			StatusBar.WriteWarning("Failed to retrieve Sales Order")
			Exit Sub
			
		End If

		oSalesOrderLines = oSalesOrder.Lines
		
		oSalesOrderLines.SetCurrentLine(pVal.Row - 1)
				
		oItmMast.GetByKey(BOMCode)
		
		oMBOM.GetByKey(BOMCode)
		
		If oMBOM.Items.Count > 0 Then
			
			
			'TEST BOM lines
			
			Dim i As Integer = 0
		
			
			oProductionOrd.ItemNo = BOMCode
			
			While i < oMBOM.Items.Count
				
				oMBOM.Items.SetCurrentLine(i)
				oProductionOrd.Lines.SetCurrentLine(i)
				oProductionOrd.Lines.ItemNo = oMBOM.Items.ItemCode
				oProductionOrd.Lines.BaseQuantity = oMBOM.Items.Quantity
				oProductionOrd.Lines.PlannedQuantity = oMBOM.Items.Quantity * oSalesOrderLines.Quantity
				oProductionOrd.Lines.ProductionOrderIssueType = oMBOM.Items.IssueMethod
				oProductionOrd.Lines.Project = oSalesOrderLines.ProjectCode
				oProductionOrd.Lines.Add
				i += 1		
				
			End While
			
			
			oProductionOrd.PlannedQuantity = oSalesOrderLines.Quantity
			oProductionOrd.CustomerCode = oSalesOrder.CardCode
			oProductionOrd.DistributionRule = oSalesOrderLines.COGSCostingCode
			oProductionOrd.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooSalesOrder
			oProductionOrd.ProductionOrderOriginEntry = Convert.ToInt16(InternalDocNum)
			oProductionOrd.ProductionOrderType = 0
			oProductionOrd.Project = oSalesOrderLines.ProjectCode
			oProductionOrd.Warehouse = "02"
			oProductionOrd.DueDate = oSalesOrderLines.ShipDate
			oProductionOrd.RoutingDateCalculation = oProductionOrd.RoutingDateCalculation.raEndDateBackwards
			Dim PrdOrdNum As Long = -1
			PrdOrdNum = oProductionOrd.Add
			
			Dim nErr As Long
			Dim errMsg As String
			
			oCompany.GetLastError(nErr, errMsg)	
			
			
			If nErr = 0 Then
					
				
				oSalesOrderLines.UserFields.Fields.Item("U_QYC_PONR").Value = oCompany.GetNewObjectKey()

				oProductionOrd = Nothing
				RunningTotal = RunningTotal + 1
				StatusBar.WriteWarning("Production Order created for: " + BOMCode + " " + QTYStr + "Pc's")
				ps.DoQuery(SubBOMQuery)
				
				If ps.RecordCount > 0 Then
					

					ps.MoveFirst
					
					BOMCode = ps.Fields.Item(0).Value.ToString
					QTYStr = ps.Fields.Item(1).Value.ToString
						
						

					While Not ps.EoF
						
						
						
						BOMCode = ps.Fields.Item(0).Value.ToString
						QTYStr = ps.Fields.Item(1).Value.ToString

						If oMBOM.GetByKey(BOMCode) = False Then
							
							StatusBar.WriteWarning("Failed to get a BOM Code: " + BOMCode)
							Exit Sub
							
						End If
						
						oProductionOrd = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
						
						
						If oMBOM.Items.Count > 0 Then
							
							oProductionOrd.ItemNo = BOMCode
							
							Dim p As Integer = 0
							
							While p < oMBOM.Items.Count
				
								oMBOM.Items.SetCurrentLine(p)
								oProductionOrd.Lines.SetCurrentLine(p)
								oProductionOrd.Lines.ItemNo = oMBOM.Items.ItemCode
								oProductionOrd.Lines.BaseQuantity = oMBOM.Items.Quantity
								oProductionOrd.Lines.PlannedQuantity = oMBOM.Items.Quantity * ps.Fields.Item(1).Value
								oProductionOrd.Lines.ProductionOrderIssueType = oMBOM.Items.IssueMethod
								oProductionOrd.Lines.Project = oSalesOrderLines.ProjectCode
								oProductionOrd.Lines.Add
								p += 1		
				
							End While
							
							
							oProductionOrd.PlannedQuantity = ps.Fields.Item(1).Value
							oProductionOrd.CustomerCode = oSalesOrder.CardCode
							oProductionOrd.DistributionRule = oSalesOrderLines.COGSCostingCode
							oProductionOrd.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooSalesOrder
							oProductionOrd.ProductionOrderOriginEntry = Convert.ToInt16(InternalDocNum)
							oProductionOrd.ProductionOrderType = 0
							oProductionOrd.Project = oSalesOrderLines.ProjectCode
							oProductionOrd.Warehouse = "02"
							oProductionOrd.DueDate = oSalesOrderLines.ShipDate
							oProductionOrd.RoutingDateCalculation = oProductionOrd.RoutingDateCalculation.raEndDateBackwards
							oProductionOrd.Add
							
							oCompany.GetLastError(nErr, errMsg)	

							If nErr = 0 Then
								
								StatusBar.WriteWarning("Production Order created for: " + BOMCode + " " + QTYStr + "Pc's Production order Number : " + oCompany.GetNewObjectKey())
								RunningTotal = RunningTotal + 1
								
							End If
						End If
										
						oProductionOrd = Nothing
						ps.MoveNext
						'i += 1
						
					End While

				End If
				
			Else
				
				StatusBar.WriteWarning("Could not create Production order for: " + BOMCode + " " + String.Format("D", QTYStr) + "Pc's")
				StatusBar.WriteWarning("Error Nbbr: " & nErr & " Error Message: " + errMsg)
				
					
			End If
			
			
		End If
		
		

		
	Else
		
		StatusBar.WriteWarning("No Line Selected")
		Exit Sub
		
	End If
	
	
	oSalesOrder.Update
	'frm.Freeze(False)
	
	frm.Refresh
	
	StatusBar.WriteSucess("Production Order creation completed. A total of " & RunningTotal & " Production order(s) were created")
	' Your Code
Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	MessageBox.Show(errorMessage, "OK")
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try
End Sub
' Function: Run_FB_COR_CUS_00000039 ----------
Public Sub Run_FB_COR_CUS_00000039(pVal As COR_Utility.Logic.CustomizeEvent)


Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim company As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
Dim oProductionOrd As SAPbobsCOM.ProductionOrders = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
Dim oPurReq As SAPbobsCOM.Documents
Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
Dim frm As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)
Dim PurReqFrm As SAPbouiCOM.Form
Dim oItem As SAPbouiCOM.Item = frm.Items.Item("37")
Dim pItem As SAPbouiCOM.Item
Dim oStatus As SAPbouiCOM.Item = frm.Items.Item("10")
Dim oRemkrs As SAPbouiCOM.Item = frm.Items.Item("3")
Dim oUpdateBtn As SAPbouiCOM.Item = frm.Items.Item("1")
Dim DistRule As String = frm.Items.Item("10000147").Specific.Value
Dim ProjCode As String = frm.Items.Item("540000153").Specific.Value
Dim mtx As SAPbouiCOM.Matrix = oItem.Specific
Dim Pmtx As SAPbouiCOM.Matrix
Dim ChkBxBomVerified As SAPbouiCOM.CheckBox = frm.Items.Item("BOYX_8").Specific
Dim ChkBxKittingComp As SAPbouiCOM.CheckBox = frm.Items.Item("BOYX_11").Specific
Dim ChkBxDecantComp As SAPbouiCOM.CheckBox = frm.Items.Item("BOYX_14").Specific
Dim ChkBx1StInspComp As SAPbouiCOM.CheckBox = frm.Items.Item("BOYX_20").Specific


Dim DecimalSep As String = System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator
 
Try
	
	Dim MenuAction As SAPbouiCOM.MenuItem = application.Menus.Item("39724")
	Dim DocEntry As String = TextEdit.GetFromUID(pVal.Form, "18").Value.ToString
	Dim PrdOrdDueDate As SAPbouiCOM.EditText = frm.Items.Item("26").Specific
	Dim Col As SAPbouiCOM.Column = mtx.Columns.Item(3)
	Dim UCol As SAPbouiCOM.Column = mtx.Columns.Item(131)
	Dim PldQTYCol As SAPbouiCOM.Column = mtx.Columns.Item(8)
	Dim EndDateCol As SAPbouiCOM.Column = mtx.Columns.Item("2340000045")
	Dim PurReqProcReason As SAPbouiCOM.ComboBox
	Dim PurReqPrdNum As SAPbouiCOM.EditText
	Dim PurReqRequireDate As SAPbouiCOM.EditText
	Dim FrmPane As SAPbouiCOM.Item
	Dim TrfUpdated As Boolean = False
	Dim DateUpdated As Boolean = False
	Dim PldQTYNum As Double = 0
	Dim Wh7QtyNum As Double = 0
	
	Dim Mode As String = frm.Mode.ToString

	
	If Not ChkBxBomVerified.Checked Then
		
		StatusBar.WriteWarning("The MPO must be Verified before you can use this Fucntion")
		Exit Sub
		
	End If
	
	If Not ChkBxKittingComp.Checked Then
		
		StatusBar.WriteWarning("The Kitting from WH01 must be completed before you can use this Fucntion")
		Exit Sub
		
	End If
	
	If Not ChkBxDecantComp.Checked Then
		
		StatusBar.WriteWarning("The Decanting must be completed before you can use this Fucntion")
		Exit Sub
		
	End If
	
	If Not ChkBx1StInspComp.Checked Then
		
		StatusBar.WriteWarning("The 1st Inspection must be completed before you can use this Fucntion")
		Exit Sub
		
	End If

	oProductionOrd.GetByKey(DocEntry)
	
	If (oProductionOrd.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned Or oProductionOrd.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased) And Mode <> "fm_UPDATE_MODE" Then
		
		MenuAction.Activate
		PurReqFrm = application.Forms.ActiveForm
		PurReqProcReason = PurReqFrm.Items.Item("NI_00010").Specific
		PurReqPrdNum = PurReqFrm.Items.Item("NI_00011").Specific
		PurReqRequireDate = PurReqFrm.Items.Item("540002106").Specific
		PurReqProcReason.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
		PurReqRequireDate.Value = PrdOrdDueDate.Value
		PurReqPrdNum.Value = DocEntry

		pItem = PurReqFrm.Items.Item("38")
	
		Pmtx = pItem.Specific

		For i As Integer = 0 To	oProductionOrd.Lines.Count - 1 Step 1
		
		
			oProductionOrd.Lines.SetCurrentLine(i)
			Dim AlreadyOrdered As Double = 0
		
			If oProductionOrd.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item And oProductionOrd.Lines.UserFields.Fields.Item("U_MTX_WHS07").Value < oProductionOrd.Lines.PlannedQuantity And oProductionOrd.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual Then
			
				
					
				Dim ExistingPONum As String = oProductionOrd.Lines.UserFields.Fields.Item("U_QYCPODON").Value.ToString
				Dim QueryStr As String = "SELECT T1.[Quantity] FROM OPOR T0 INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[ItemCode] = '" + oProductionOrd.Lines.ItemNo + "' AND  T1.[U_QYC_PONR] = '" + DocEntry + "' AND T0.[CANCELED] <> 'Y'"
									
					rs.DoQuery(QueryStr)
					
					If rs.RecordCount > 0 Then
						
						rs.MoveFirst
						AlreadyOrdered = rs.Fields.Item(0).Value
						StatusBar.WriteWarning(oProductionOrd.Lines.ItemNo + " Already Ordered: " & AlreadyOrdered)
						
					End If
					
					
				
				
				If AlreadyOrdered < oProductionOrd.PlannedQuantity - oProductionOrd.Lines.UserFields.Fields.Item("U_MTX_WHS07").Value Then 
					
					Pmtx.AddRow()
					Pmtx.Columns.Item("1").Cells.Item(Pmtx.RowCount - 1).Specific.Value = oProductionOrd.Lines.ItemNo
					Pmtx.Columns.Item("2004").Cells.Item(Pmtx.RowCount - 1).Specific.Value = DistRule
					Pmtx.Columns.Item("31").Cells.Item(Pmtx.RowCount - 1).Specific.Value = ProjCode
					Pmtx.Columns.Item("U_QYCPrdEntry").Cells.Item(Pmtx.RowCount - 1).Specific.Value = DocEntry
					Pmtx.Columns.Item("11").Cells.Item(Pmtx.RowCount - 1).Specific.Value = oProductionOrd.Lines.PlannedQuantity - oProductionOrd.Lines.UserFields.Fields.Item("U_MTX_WHS07").Value - AlreadyOrdered
					Pmtx.Columns.Item("U_QYCPrdLine").Cells.Item(Pmtx.RowCount - 1).Specific.Value = oProductionOrd.Lines.LineNumber + 1
					Pmtx.Columns.Item("24").Cells.Item(Pmtx.RowCount - 1).Specific.Value = "01"
					Pmtx.Columns.Item("540002123").Cells.Item(Pmtx.RowCount - 1).Specific.Value = PrdOrdDueDate.Value
				
				
				End If
					
			End If
		
		
		Next
		
		
	End If
	

	' Your Code
Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	MessageBox.Show(errorMessage, "OK")
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try
End Sub
' Function: Run_FB_COR_CUS_00000040 ----------
Public Sub Run_FB_COR_CUS_00000040(pVal As COR_Utility.Logic.CustomizeEvent)

 
Dim application As SAPbouiCOM.Application = customize.B1Connector.GetB1Connector().Application
Dim company As SAPbobsCOM.Company = customize.B1Connector.GetB1Connector().Company
Dim frmPrd As SAPbouiCOM.Form = application.Forms.Item(pVal.FormUID)
Dim oItem As SAPbouiCOM.Item = frmPrd.Items.Item("37")
Dim mtx As SAPbouiCOM.Matrix = oItem.Specific



Try
	
	
	
	Dim ImgCol As SAPbouiCOM.Column = mtx.Columns.Item("U_ImageKITIssuedStoc")

	
	StatusBar.WriteWarning("DEBUG - Rule: " + pVal.RuleInfo.RuleName + " was triggered.")
	
	Dim PictureStr As String = Nothing
	
	If pVal.Row = -1 Then
		
		Exit Sub
		
	End If
	
		
	PictureStr = ImgCol.Cells.Item(pVal.Row).Specific.Value.ToString
	
	If Not IsNothing(PictureStr) And PictureStr.Length > 0 Then
		
		ViewImgFromSapLIB.MyUtils.ViewImageFromSapAsync(PictureStr)

		
		
		Exit Sub

		
	Else
		
		
		Exit Sub
		
		
		
		
	End If
	
	

	
	
	
Catch ex As Exception
	Dim errorMessage As String = String.Format("Error in {0} Rule '{1}': {2}", pVal.RuleInfo.RuleType, pVal.RuleInfo.RuleName, ex.Message)
	MessageBox.Show(errorMessage, "OK")
	StatusBar.WriteError(errorMessage)
	Debug.WriteMessage(errorMessage, Debug.DebugLevel.Exception)
End Try
End Sub


End Class
End Namespace
