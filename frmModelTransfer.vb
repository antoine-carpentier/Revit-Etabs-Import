'
' Created by SharpDevelop.
' User: michael
' Date: 3/27/2015
' Time: 2:58 PM
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Imports System
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI.Selection
Imports System.Collections.Generic
Imports System.Linq
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Windows.Forms
Imports System.IO
Imports LumenWorks.Framework.IO.Csv
Imports System.Data

Public Partial Class frmRenameSections
	
	
	Public Sub New(curDoc As Document)
		' The Me.InitializeComponent call is required for Windows Forms designer support.
		Me.InitializeComponent()
		
		'
		' TODO : Add constructor code after InitializeComponents
		'				
		
		'Set up DataGrid
		'dataGridViewSection.ColumnCount=3

	End Sub
	
'	Sub button1Click(sender As Object, e As System.EventArgs) Handles button1.Click
'		
'		
'		
'		Dim CSVPath As String = CSVLocation.Text
'		
'        Dim dt As New DataTable
'
'            dim csvR As New CsvReader(New StreamReader(CSVPath), false)
'            dt.Load(csvR)           
'            	
'
'		Dim data2 As DataTable = Dt.Copy()
'
'		data2.Rows(0).Delete
'		data2.Rows(1).Delete
'		data2.Rows(2).Delete
'		
'		Dim Test As Object = data2.Compute("Max(Column1)","")
'		
'		TaskDialog.Show("test", test.ToString)
'
'
'		
'	End Sub
	
	
	
	
	Sub CSVBrowseClick(sender As Object, e As System.EventArgs) Handles CsvBrowse.Click
	
		CSVOpenFile.RestoreDirectory = True
		CSVOpenFile.Title = "Select a CSV File"
		CSVOpenFile.Filter = "CSV Files (*.csv)|*.csv"
		
		If CSVOpenFile.ShowDialog() <> System.Windows.Forms.DialogResult.Cancel Then
            CSVLocation.Text = CSVOpenFile.FileName
        End If
	
	End Sub
	
	
	'Sub OpenCSV 
	Function OpenCSV As datatable 
		
		
		Dim CVSPath As String = CSVLocation.Text
		
        Dim dt As New DataTable

            Using csvR As New CsvReader(New StreamReader(CVSPath), false)
            	dt.Load(csvR)            	
            	
            End Using

        Return dt	
		
					        
	end function
	'End Sub
	
	Public Function GridCheck(curdoc As Document) As Boolean
		
		If Me.GridcheckBox.Checked = True Then
			Return True
		Else return false
		End If
		
	End Function
	
	Public Function LevelCheck(curdoc As Document) As Boolean
		
		If Me.LevelcheckBox.Checked = True Then
			Return True
		Else return false
		End If
		
	End Function
		
	Public Function WallCheck(curdoc As Document) As Boolean
		
		If Me.WallcheckBox.Checked = True Then
			Return True
		Else return false
		End If
		
	End Function
			
	Public Function SlabCheck(curdoc As Document) As Boolean
		
		If Me.SlabcheckBox.Checked = True Then
			Return True
		Else return false
		End If
		
	End Function
	
	Public Function ColCheck(curdoc As Document) As Boolean
		
		If Me.ColcheckBox.Checked = True Then
			Return True
		Else return false
		End If
		
	End Function
	
	Public Function BeamCheck(curdoc As Document) As Boolean
		
		If Me.BeamcheckBox.Checked = True Then
			Return True
		Else return false
		End If
		
	End Function
	
End Class