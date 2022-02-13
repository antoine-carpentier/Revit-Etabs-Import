
' Created by SharpDevelop.
' User: acarpentier
' Date: 12/21/2018
' Time: 11:30 AM
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
Imports System.Data
Imports Microsoft.VisualBasic

<Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)> _
<Autodesk.Revit.DB.Macros.AddInId("7BD4FE06-702E-4A54-88B0-A70EFD65F60B")> _
Partial Public Class ThisDocument
	
	Public StartTime As DateTime=DateTime.Now
	
	Public CSVData As New DataTable
	Public WarningError As String =""
	
	Public GridWarningError As String =""
	Public LevelWarningError As String =""
	Public BeamWarningError As String =""
	Public ColWarningError As String =""
	Public WallWarningError As String =""
	Public SlabWarningError As String =""
	
	Public LevelBool As Boolean = False
	Public GridBool As Boolean = False
	Public WallBool As Boolean = False
	Public SlabBool As Boolean = False
	Public ColBool As Boolean = False
	Public BeamBool As Boolean = False
	
	
	Public ProblemColId As List(Of ElementId)
	Public ColIdMatch As New List(Of List(Of Object))
	
	
	
	
	Public Sub ModelTransfer()
		
		'Get the current document
		Dim curDoc As Document = Me.Application.ActiveUIDocument.Document


				Using curForm As New frmRenameSections(curDoc)
					
					'show form
					curForm.ShowDialog
					
					If curForm.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
						'Result.Cancelled
						
					Else
						
						WarningError = ""
						GridWarningError = ""
						LevelWarningError = ""
						BeamWarningError = ""
						ColWarningError = ""
						WallWarningError = ""
						SlabWarningError = ""
						
						LevelBool = curForm.LevelCheck(curDoc)
						GridBool = curForm.GridCheck(curDoc)
						SlabBool = curForm.SlabCheck(curDoc)
						WallBool = curForm.WallCheck(curDoc)
						ColBool = curform.ColCheck(curDoc)
						BeamBool = curForm.BeamCheck(curDoc)
						
						
						
						CSVData = curForm.OpenCSV()
						
						'Start transaction group
						
						'start transaction
						Using curTrans As New Transaction(curDoc, "Model Import")
							
							'Initialize the failure preprocessor
							Dim failureHandlingOptions As FailureHandlingOptions = curTrans.GetFailureHandlingOptions()
							Dim FailureHandlerlocal As FailureHandler = New FailureHandler	
							failureHandlingOptions.SetFailuresPreprocessor(FailureHandlerlocal)		
							failureHandlingOptions.SetClearAfterRollback(True)	
							curTrans.SetFailureHandlingOptions(failureHandlingOptions)	
							
							'Get the list of columns that are problematic
							ProblemColId = FailureHandlerlocal.ProblemColId	
							
							
							If curTrans.Start = TransactionStatus.Started Then						
								
								MainSub(curDoc)	
					
							End If 
								
									
							'commit changes
							curTrans.Commit
							
						End Using
					
						If ProblemColId.Count <>0 Then		
							
							Using curTrans As New Transaction(curDoc, "Model Import")
							
'								'Initialize the failure preprocessor
'								Dim failureHandlingOptions As FailureHandlingOptions = curTrans.GetFailureHandlingOptions()
'								Dim FailureHandlerlocal As FailureHandler = New FailureHandler	
'								failureHandlingOptions.SetFailuresPreprocessor(FailureHandlerlocal)		
'								failureHandlingOptions.SetClearAfterRollback(False)	
'								curTrans.SetFailureHandlingOptions(failureHandlingOptions)	
'								
'								'Get the list of columns that are problematic
'								ProblemColId = FailureHandlerlocal.ProblemColId	
								
								
								If curTrans.Start = TransactionStatus.Started Then						
									
									MainSub(curDoc)	
						
								End If 
									
										
								'commit changes
								curTrans.Commit
							
							End Using
							
						End If


						Dim EndTime As DateTime = DateTime.Now
						Dim TimeElapsed As TimeSpan = endTime-startTime

						WarningError = WarningError & vbCrLf & "Time elapsed = " & TimeElapsed.Minutes & " minutes."

						TaskDialog.Show("Model Import", WarningError)
						
					End If 
			End Using
 
	End Sub
	
	Public Sub MainSub(Curdoc As Document)
		
		WarningError = ""
		GridWarningError = ""
		LevelWarningError = ""
		BeamWarningError = ""
		ColWarningError = ""
		WallWarningError = ""
		SlabWarningError = ""
		
			Using curForm As New frmRenameSections(curDoc)
				
				If LevelBool = True Then
					Try
						CreateLevels(curDoc)
					Catch ex As Exception
						LevelWarningError = "     - ERROR - Something unexpected happened. The levels couldn't be created."
					End Try				
				End if
				
				
				If GridBool = True Then	
					Try
						CreateGrids(curDoc)	
					Catch ex As Exception
						GridWarningError = "     - ERROR - Something unexpected happened. The grids couldn't be created."
					End Try
				end if
				
				
				If ColBool = True Then
					Try
						GetColumns(curDoc)
					Catch ex As Exception
						ColWarningError = "     - ERROR - Something unexpected happened. The columns couldn't be created."
					End Try
				End If
				

				If BeamBool = True Then
					Try
						GetBeams(curDoc)
					Catch ex As Exception
						BeamWarningError = "     - ERROR - Something unexpected happened. The beams couldn't be created."
					End Try
				End If
								
								
				If SlabBool = True Then
					Try
						GetSlabs(curDoc)
					Catch ex As Exception
						SlabWarningError = "     - ERROR - Something unexpected happened. The slabs couldn't be created."
					End Try
				End If
								

				If WallBool = True Then
					Try
						GetWalls(curDoc)
					Catch ex As Exception
						WallWarningError = "     - ERROR - Something unexpected happened. The walls couldn't be created."
					End Try
				End If

						If LevelWarningError = "" Then
							LevelWarningError = "     - No errors or warnings."  & vbCrLf
						End If
						
						
						If GridWarningError = "" Then
							GridWarningError = "     - No errors or warnings."  & vbCrLf
						End if
						
						If BeamWarningError = "" Then
							BeamWarningError = "     - No errors or warnings."  & vbCrLf
						End If
						
						If ColWarningError = "" Then
							ColWarningError = "     - No errors or warnings."  & vbCrLf
						End If
						
						If SlabWarningError = "" Then
							SlabWarningError = "     - No errors or warnings."  & vbCrLf
						End If
						
						If WallWarningError = "" Then
							WallWarningError = "     - No errors or warnings." & vbCrLf
						End If
						
						If LevelBool = True Then
							WarningError = WarningError & vbCrLf & "• Levels:" & vbCrLf & LevelWarningError 
						End If
						
						If gridbool = True Then
							WarningError = WarningError & vbCrLf & "• Grids:" & vbCrLf & GridWarningError
						End If
						
						If BeamBool = True Then
							WarningError = WarningError & vbCrLf & "• Beams:" & vbCrLf & beamWarningError
						End If
						
						If ColBool = True Then
							WarningError = WarningError & vbCrLf & "• Columns:" & vbCrLf & colWarningError
						End If
						If SlabBool = True Then
							WarningError = WarningError & vbCrLf & "• Slabs:" & vbCrLf & slabWarningError
						End If
						If WallBool = True Then
							WarningError = WarningError & vbCrLf & "• Walls:" & vbCrLf & wallWarningError
						End if
								
			End using
		
	End Sub
	
	Public Sub CreateLevels(CurDoc As Document)
		
		'create all levels
		
		Dim collector As New clsCollectors
		Dim LevelList As List(Of Level) = collector.GetAllLevels(CurDoc)
		
		Dim LevelElevationList As New List(Of Double)
		For Each lvl As Level In LevelList
			LevelElevationList.Add(lvl.Elevation)
		Next
		
								For i=0 To CSVData.Rows(1).Item(12).ToString-1
									If LevelElevationList.Contains(CSVData.Rows(3+i).Item(13).ToString) Then
										LevelWarningError = LevelWarningError & "     - There is already a level with the same elevation as """ & CSVData.Rows(3+i).Item(12).ToString & """. This level won't be created." & vbCrLf
'										TaskDialog.Show("Level already exists", "There is already a level with the same elevation as """ & CSVData.Rows(3+i).Item(12).ToString & """. This level won't be created.")
									Else
'										
										Dim NewLevel As Level = Level.Create(CurDoc,CSVData.Rows(3+i).Item(13).ToString)
										
   										If NewLevel Is Nothing Then
    										Throw New Exception("Create a new level failed.")
   										End If
   	
   	
   										' Modify the name of the created grid
   										Try 
   											newlevel.Name = CSVData.Rows(3+i).Item(12).ToString
   										Exit Try
   										Catch e As Autodesk.Revit.Exceptions.ArgumentException
   											LevelWarningError = LevelWarningError & "     - " & CSVData.Rows(3+i).Item(12).ToString & " is already used to name a different level. " & NewLevel.Name & " will be used instead." & vbCrLf
'   											TaskDialog.Show("Error",CSVData.Rows(3+i).Item(12).ToString & " is already used to name a different level. " & NewLevel.Name & " will be used instead.")
   										End Try
									End If
								Next
		
		
	End Sub
	
	Public Sub CreateGrids(CurDoc As Document)
		
		Dim collector As New clsCollectors
		Dim Gridlist As List(Of Grid) = collector.GetAllGrids(CurDoc)
		Dim StartPoint As New List (Of XYZ)
		Dim EndPoint As New List (Of XYZ)
		Dim ExistXGrid As New List (Of Double)
		Dim ExistYGrid As New List (Of Double)
		Dim ExistGridEquation1 As New List(Of Double)
		Dim ExistGridEquation2 As New List(Of Double)
		Dim ExistGridEquation3 As New List(Of Double)
		
		'Get all existing grid to see if we are duplicating some
		
		
		For Each CurGrid As Grid In Gridlist
			Dim Curcurve As Curve = CurGrid.Curve
			
			'List the ▲ X Grids
			If Curcurve.GetEndPoint(0).X= Curcurve.GetEndPoint(1).X Then
				ExistXGrid.Add(Curcurve.GetEndPoint(1).X)	
				
			'List the ▲ Y Grids
			ElseIf Curcurve.GetEndPoint(0).y= Curcurve.GetEndPoint(1).y Then
				ExistyGrid.Add(Curcurve.GetEndPoint(1).y)

			Else 
				'Get the ax+by+c=0 equation of all gen grids
				
				ExistGridEquation1.Add(Math.Round(Curcurve.GetEndPoint(0).y-Curcurve.GetEndPoint(1).Y,3))
				ExistGridEquation2.Add(Math.Round(Curcurve.GetEndPoint(1).x-Curcurve.GetEndPoint(0).x,3))
				ExistGridEquation3.Add(Math.Round(Curcurve.GetEndPoint(0).x*Curcurve.GetEndPoint(1).y-Curcurve.GetEndPoint(1).x*Curcurve.GetEndPoint(0).y,3))
 
			End If

		Next
		
		
		Dim MinX As Double = -100
		Dim maxX As Double = 100
		Dim MinY As Double = -100
		Dim maxY As Double = 100
		
		If CSVData.Rows(1).Item(0) > 0 Then
			MinX = CSVData.Rows(3).Item(1).ToString - 5
			MaxX = CSVData.Rows(CSVData.Rows(1).Item(0).ToString + 2).Item(1).ToString + 5
		End if
	
		
		If CSVData.Rows(1).Item(3) > 0 Then
			MinY = CSVData.Rows(3).Item(4).ToString - 5
			MaxY = CSVData.Rows(CSVData.Rows(1).Item(3).ToString + 2).Item(4).ToString + 5
		End if
			
								
								
								'create all x grids
								
								For i=0 To CSVData.Rows(1).Item(0).ToString-1
									
									If ExistXGrid.Contains(CSVData.Rows(3+i).Item(1).ToString) Then
										GridWarningError = GridWarningError & "     - There is already a grid matching " & CSVData.Rows(3+i).Item(0).ToString & "." & vbCrLf
'										TaskDialog.Show("Grid already exists", "There is already a grid matching " & CSVData.Rows(3+i).Item(0).ToString & ".")
									Else  
											Dim StartPt As New XYZ(CSVData.Rows(3+i).Item(1).ToString,MinY,0)
											Dim EndPt As New xyz(CSVData.Rows(3+i).Item(1).ToString,MaxY,0)
											Dim GeoLine As Line = Line.CreateBound(StartPt,EndPt)
	
									' Create a grid using the geometry line
	
									Dim lineGrid As Grid = Grid.Create(CurDoc,GeoLine)
	
   									If lineGrid Is Nothing Then
   										Throw New Exception("Create a new straight grid failed.")
   										Else gridlist.Add(lineGrid)
   									End If
   	
   	
   									' Modify the name of the created grid

   									Try 
   										lineGrid.Name = CSVData.Rows(3+i).Item(0).ToString
   										Exit Try
   									Catch e As Autodesk.Revit.Exceptions.ArgumentException
   										GridWarningError = GridWarningError & "     - " & CSVData.Rows(3+i).Item(0).ToString & " is already used to name a different grid. " & lineGrid.Name & " will be used instead." & vbCrLf
'   										TaskDialog.Show("Error",CSVData.Rows(3+i).Item(0).ToString & " is already used to name a different grid. " & lineGrid.Name & " will be used instead.")
   									End Try
   									
   									End If
   									
								Next
								

								
								'create all y grids
								
								For i=0 To CSVData.Rows(1).Item(3).ToString-1
									
									If ExistYGrid.Contains(CSVData.Rows(3+i).Item(4).ToString) Then
										GridWarningError = GridWarningError & "     - There is already a grid matching " & CSVData.Rows(3+i).Item(3).ToString & "." & vbCrLf
'										TaskDialog.Show("Grid already exists", "There is already a grid matching " & CSVData.Rows(3+i).Item(3).ToString & ".")
									Else  
								
									Dim StartPt As New XYZ(MinX,CSVData.Rows(3+i).Item(4).ToString,0)
									Dim EndPt As New xyz(MaxX,CSVData.Rows(3+i).Item(4).ToString,0)
									Dim GeoLine As Line = Line.CreateBound(StartPt,EndPt)
	
									' Create a grid using the geometry line
	
									Dim lineGrid As Grid = Grid.Create(CurDoc,GeoLine)
	
	
   									If lineGrid Is Nothing Then
   										Throw New Exception("Create a new straight grid failed.")
   										Else gridlist.Add(lineGrid)
   									End If
   	
   	
   									' Modify the name of the created grid

   									Try 
   										lineGrid.Name = CSVData.Rows(3+i).Item(3).ToString
   										Exit Try
   									Catch e As Autodesk.Revit.Exceptions.ArgumentException
   										GridWarningError = GridWarningError & "     - " & CSVData.Rows(3+i).Item(3).ToString & " is already used to name a different grid. " & lineGrid.Name & " will be used instead." & vbCrLf
'   										TaskDialog.Show("Error",CSVData.Rows(3+i).Item(3).ToString & " is already used to name a different grid. " & lineGrid.Name & " will be used instead.")
   									End Try
   									
   									End If 
								Next
								
								
								
								'create all gen grids
								
								
								
								For i=0 To CSVData.Rows(1).Item(6).ToString-1
									
									
									Dim CurStartX As  Double = CSVData.Rows(3+i).Item(7)
									Dim CurStartY As  Double = CSVData.Rows(3+i).Item(8)
									Dim CurEndX As Double = CSVData.Rows(3+i).Item(9)
									Dim CurEndY As Double = CSVData.Rows(3+i).Item(10)
									
									Dim cureq1 As Double = Math.Round(CurStartY-CurEndY,3)
									Dim cureq2 As  Double = Math.Round(CurEndX-CurStartX,3)
									Dim cureq3 As Double = Math.Round(CurStartX*CurEndY-CurEndX*CurStartY,3)
								
									
									If (ExistGridEquation1.Contains(cureq1) And ExistGridEquation2.Contains(cureq2)  And ExistGridEquation3.Contains(cureq3)) Or _
										(ExistGridEquation1.Contains(-cureq1) And ExistGridEquation2.Contains(-cureq2)  And ExistGridEquation3.Contains(-cureq3)) Then
										
										GridWarningError = GridWarningError & "     - There is already a grid matching " & CSVData.Rows(3+i).Item(6).ToString & "." & vbCrLf
'										TaskDialog.Show("Grid already exists", "There is already a grid matching " & CSVData.Rows(3+i).Item(6).ToString & ".")

									Else 
									
																		
									Dim StartPt As New XYZ(CSVData.Rows(3+i).Item(7).ToString,CSVData.Rows(3+i).Item(8).ToString,0)
									Dim EndPt As New xyz(CSVData.Rows(3+i).Item(9).ToString,CSVData.Rows(3+i).Item(10).ToString,0)
									Dim GeoLine As Line = Line.CreateBound(StartPt,EndPt)
	
									' Create a grid using the geometry line
	
									Dim lineGrid As Grid = Grid.Create(CurDoc,GeoLine)
	
	
   									If lineGrid Is Nothing Then
   										Throw New Exception("Create a new straight grid failed.")
   										Else gridlist.Add(lineGrid)
   									End If
   	
   	
   									' Modify the name of the created grid
   									Try 
   										lineGrid.Name = CSVData.Rows(3+i).Item(6).ToString
   										Exit Try
   									Catch e As Autodesk.Revit.Exceptions.ArgumentException
   										GridWarningError = GridWarningError & "     - " & CSVData.Rows(3+i).Item(6).ToString & " is already used to name a different grid. " & lineGrid.Name & " will be used instead." & vbCrLf
'   										TaskDialog.Show("Error",CSVData.Rows(3+i).Item(6).ToString & " is already used to name a different grid. " & lineGrid.Name & " will be used instead.")
   									End Try
   										
									End If
								Next

								
	End Sub
	
	Public Sub GetColumns(CurDoc As Document)
				
				
		Dim ColList As New List(Of list(Of String))  'List of X1,Y1,Z1,X2,Y2,Z2,shape
		Dim NewList As New List(Of String)
		Dim ShapeList As New List(Of String)
		
		Dim VersionPath As String = ""
		
		'Check for Revit version		
		If Application.Application.VersionNumber = "2015" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2015\"
		ElseIf Application.Application.VersionNumber = "2016" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2016\"
		ElseIf Application.Application.VersionNumber = "2017" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2017\"
		ElseIf Application.Application.VersionNumber = "2018" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2018\"
		ElseIf Application.Application.VersionNumber = "2019" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2019\"
		End If
		
		'Get all already loaded family and their names in separate lists
		Dim collector As New clsCollectors
		Dim LoadedColList As List(Of Family) = collector.getAllColumnFamilies(CurDoc)
		Dim LoadedColListName As New List(Of String)
		
		For Each Fam As Family In LoadedColList
			LoadedColListName.Add(Fam.Name)
		Next
		
		'Make sure that at least one round and one rectangular column is loaded, and assign templates
		Dim RectangularTemplate As FamilySymbol = Nothing
		Dim RoundTemplate As FamilySymbol = Nothing
		Dim AllRectangular As New List(Of FamilySymbol)
		Dim AllRound As New List(Of FamilySymbol)
		Dim tempList As New HashSet(Of ElementId) 'List(Of ElementId)
		
		If LoadedColListName.Contains("Concrete-Rectangular-Column") Then
			For Each fam As Family In LoadedColList
				If fam.Name = "Concrete-Rectangular-Column" Then
					
					tempList = fam.GetFamilySymbolIds
					
					For Each elid As ElementId In tempList
						AllRectangular.Add(DirectCast(CurDoc.GetElement(elid),FamilySymbol))
					Next
					
					RectangularTemplate = DirectCast(CurDoc.GetElement(fam.GetFamilySymbolIds(0)),FamilySymbol)					
				End If
			Next
		Else curdoc.LoadFamilySymbol(VersionPath & "Libraries\US Imperial\Structural Columns\Concrete\Concrete-Rectangular-Column.rfa","12 x 18",New FamilyLoadOptions(),RectangularTemplate)
		End If
		
		tempList = New HashSet(Of ElementId)
		
		If LoadedColListName.Contains("Concrete-Round-Column") Then
			For Each fam As Family In LoadedColList
				If fam.Name = "Concrete-Round-Column" Then
					
					tempList = fam.GetFamilySymbolIds
					
					For Each elid As ElementId In tempList
						Allround.Add(DirectCast(CurDoc.GetElement(elid),FamilySymbol))
					Next
					
					RoundTemplate = DirectCast(CurDoc.GetElement(fam.GetFamilySymbolIds(0)),FamilySymbol)					
				End If
			Next
		Else curdoc.LoadFamilySymbol(VersionPath & "Libraries\US Imperial\Structural Columns\Concrete\Concrete-Round-Column.rfa","12""",New FamilyLoadOptions(),RoundTemplate)
		End If
		

		
		'Store the col CSV file data in memory, and create a separate list containing only the shapes		
		For i=0 To CSVData.Rows(1).Item(35).ToString-1
			If CSVData.Rows(3+i).Item(42) <45 Then
				NewList = New List(Of String)
				NewList.Add(CSVData.Rows(3+i).Item(36).ToString)   '
				NewList.Add(CSVData.Rows(3+i).Item(37).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(38).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(39).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(40).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(41).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(43).ToString)
				ShapeList.Add(CSVData.Rows(3+i).Item(43).ToString)
				ColList.Add(NewList)
			End if
		Next
	
		
		'Eliminate duplicates from the ShapeList		
		ShapeList = ShapeList.Distinct().ToList
		
		'get dimensions of shapes that are not steel		
		Dim ShapeDims As New List(Of List(Of String))
		
		For i=0 To CSVData.Rows(1).Item(20).ToString-1
			If ShapeList.Contains(CSVData.Rows(3+i).Item(20).ToString) Then
				NewList = New List(Of String)
				NewList.add(CSVData.Rows(3+i).Item(20).ToString)
				NewList.add(CSVData.Rows(3+i).Item(22).ToString)
				NewList.add(CSVData.Rows(3+i).Item(23).ToString)
				ShapeDims.Add(NewList)
			End If
		Next
		
		
		
		'For each shape used, check what type of shape it is
		' list of possible types: I-shape / Channel / T / Angle / DblAngle / Box / Pipe / Rectangular / Circle / DblChannel / Other 
		
		Dim TypeList As New List(Of String)
		
		For Each CurShape As String In ShapeList
			For i=0 To CSVData.Rows(1).Item(20).ToString-1
				If CSVData.Rows(3+i).Item(20).ToString = CurShape Then
					TypeList.Add(CSVData.Rows(3+i).Item(21).ToString)
				End If
			Next
		Next
		
		
		'Associate each Shape with its Revit symbol
		
		Dim NameSymbolPairs As New List(Of List(Of Object))
		Dim ShapePath As String = ""
		Dim VarList As New List(Of object)
		Dim CurSymbol As FamilySymbol = Nothing
		Dim b As Double = 0
		Dim h As Double = 0
		Dim Loaded As Boolean = False
		Dim IdentificationCounter As Integer = 1
		
		'Create an template steel shape for poorly named shapes in CSV file\
			Dim BaseShape As FamilySymbol = Nothing
			CurDoc.LoadFamilySymbol(VersionPath & "Libraries\US Imperial\Structural Columns\Steel\W-Wide Flange-Column.rfa","W12X26",New FamilyLoadOptions(), BaseShape)

		For i=0 To TypeList.Count-1 
			
			VarList = New  List(Of Object)

			'If it's a steel shape
			If TypeList(i)= "I-Shape" Or TypeList(i)= "Channel" Or TypeList(i)= "Angle" Or TypeList(i)= "DblAngle" Or TypeList(i)= "Box" Or TypeList(i)= "Pipe" Or TypeList(i)= "DblChannel" Then
				
				If TypeList(i)= "I-Shape" Then
					ShapePath = "Libraries\US Imperial\Structural Columns\Steel\W-Wide Flange-Column.rfa"
				ElseIf TypeList(i)= "Channel" Then
					ShapePath = "Libraries\US Imperial\Structural Columns\Steel\C-Channel-Column.rfa"
				ElseIf TypeList(i)= "Angle" Then
					ShapePath = "Libraries\US Imperial\Structural Columns\Steel\L-Angle-Column.rfa"
'				ElseIf TypeList(i)= "DblAngle" Then
'					ShapePath = "Libraries\US Imperial\Structural Columns\Steel\W-Wide Flange-Column.rfa"
				ElseIf TypeList(i)= "Box" Then
					ShapePath = "Libraries\US Imperial\Structural Columns\Steel\HSS-Hollow Structural Section-Column.rfa"
				ElseIf TypeList(i)= "Pipe" Then
					ShapePath = "Libraries\US Imperial\Structural Columns\Steel\HSS-Round Hollow Structural Section-Column.rfa"
				ElseIf TypeList(i)= "DblChannel" Then
					ShapePath = "Libraries\US Imperial\Structural Columns\Steel\Double C-Channel-Column.rfa"
				End If
				
				Dim CurName As String =""
				
				If ShapeList(i).EndsWith("-1") Then
					CurName = ShapeList(i).Substring(0,ShapeList(i).Length-2)
				Else
					CurName = ShapeList(i)
				End If
				
				Loaded = CurDoc.LoadFamilySymbol(VersionPath & ShapePath,CurName, CurSymbol)
				
				If Loaded = False Then
					CurSymbol = BaseShape.Duplicate(CSVData.Rows(0).Item(0).ToString & " - UNIDENTIFIED " & IdentificationCounter)
					ColWarningError = ColWarningError & "     - """ & shapeList(i) & """ is not a valid Revit shape. """ & CurSymbol.Name & """ will be used instead. " & vbCrLf 
					IdentificationCounter = IdentificationCounter+1
				End If
				
				CurSymbol.Activate()
				
				VarList.Add(ShapeList(i))
				VarList.Add(CurSymbol)
				
				NameSymbolPairs.Add(VarList)
				
				
				
				
			'If it's a concrete or similar material, then	
			ElseIf TypeList(i)= "Rectangular" Or TypeList(i)= "Circle" Or TypeList(i)= "Other" 	Then
				
				Dim FamSymExist As Boolean = False
				
				
				If TypeList(i)= "Rectangular" Or TypeList(i)= "Other" Then
					
					Dim ColName As String = CSVData.Rows(0).Item(0).ToString & " - " & ShapeList(i)
					
					'Deleted prohibited character from string
					If ColName.Contains(":") Or ColName.Contains(";") Or ColName.Contains("<") Or ColName.Contains(">") Or ColName.Contains("[") Or ColName.Contains("]") Then
						ColName = ColName.Replace(":","")
						ColName = ColName.Replace(";","")
						ColName = ColName.Replace("<","")
						ColName = ColName.Replace(">","")
						ColName = ColName.Replace("[","")
						ColName = ColName.Replace("]","")			
					End If
					
					For Each FamSym As FamilySymbol In AllRectangular
						If FamSym.Name = ColName Then
							FamSymExist = True
							VarList.Add(ShapeList(i))
							VarList.Add(famSym)
				
							NameSymbolPairs.Add(VarList)
						End If
					Next
					
					If FamSymExist = False Then
						CurSymbol= RectangularTemplate.Duplicate(ColName)
						
						For u=0 To ShapeDims.Count-1
							If ShapeDims(u)(0) = ShapeList(i) Then
								b = ShapeDims(u)(1)/12
								h = ShapeDims(u)(2)/12
							End If
						Next
						
						CurSymbol.Activate()
						
						CurSymbol.GetOrderedParameters(1).Set(b)
						CurSymbol.GetOrderedParameters(2).Set(h)
							
						VarList.Add(ShapeList(i))
						VarList.Add(CurSymbol)
				
						NameSymbolPairs.Add(VarList)
								
					End If		
					
				ElseIf  TypeList(i)= "Circle" Then
							
					Dim ColName As String = CSVData.Rows(0).Item(0).ToString & " - " & ShapeList(i)
					
					'Deleted prohibited character from string
					If ColName.Contains(":") Or ColName.Contains(";") Or ColName.Contains("<") Or ColName.Contains(">") Or ColName.Contains("[") Or ColName.Contains("]") Then
						ColName = ColName.Replace(":","")
						ColName = ColName.Replace(";","")
						ColName = ColName.Replace("<","")
						ColName = ColName.Replace(">","")
						ColName = ColName.Replace("[","")
						ColName = ColName.Replace("]","")			
					End If
					
					For Each FamSym As FamilySymbol In AllRound
						If FamSym.Name = ColName Then
							FamSymExist = True
							VarList.Add(ShapeList(i))
							VarList.Add(famSym)
				
							NameSymbolPairs.Add(VarList)
						End If
					Next
					
					If FamSymExist = False Then
						CurSymbol= RoundTemplate.Duplicate(ColName)
						
						For u=0 To ShapeDims.Count-1
							If ShapeDims(u)(0) = ShapeList(i) Then
								b = ShapeDims(u)(1)/12
							End If
						Next
						
						CurSymbol.GetOrderedParameters(1).Set(b)
							
						VarList.Add(ShapeList(i))
						VarList.Add(CurSymbol)
				
						NameSymbolPairs.Add(VarList)
								
					End If
				End If
				
			End If

		Next

	
	Dim LevelList As List(Of Level) = collector.GetAllLevels(CurDoc)

		
		'Place all the columns
		Dim Newlist5 As New List(Of Object)
		Dim SkipBool As Boolean = false
		
		For i=0 To ColList.Count-1
			
			SkipBool=False
			
			For u =0 To ProblemColId.Count-1
				If ColIdMatch(i)(1) = ProblemColId(u) Then
					SkipBool = True
'					TaskDialog.Show("P", "A Match!")
				End If
			Next
			
			If SkipBool= False then
				
				Newlist5=New List(Of Object)
	
				Dim CurStart As New XYZ(ColList(i)(0),ColList(i)(1),ColList(i)(2))
				Dim CurEnd As New XYZ(ColList(i)(3),ColList(i)(4),ColList(i)(5))
				Dim CurLine As Curve = Line.CreateBound(CurStart,CurEnd)
	
				Dim Curshape As FamilySymbol = Nothing
	
				For u=0 To NameSymbolPairs.Count-1
					If NameSymbolPairs(u)(0) = ColList(i)(6) Then
						CurShape=NameSymbolPairs(u)(1)
					End If
				Next
					
				Dim curlevel As Level = LevelList(0)
				Dim LevelOffset As Double = 999
				Dim CurLowPoint As Double = Math.Min(CDbl(ColList(i)(2)),CDbl(ColList(i)(5)))
				Dim Curcol As FamilyInstance = Nothing
					
				For Each lvl As Level In LevelList
					If Math.Abs(lvl.Elevation-CurLowPoint)<LevelOffset Then
						curlevel = lvl
						LevelOffset = Math.Abs(lvl.Elevation-CurLowPoint)
					End If
				Next
				
				If Curshape Is Nothing Then
	''				ColWarningError = ColWarningError & "     - """ & ColList(i)(6).ToString & """ is not a valid Revit shape (at CSV line " & i+4 & "). ""ETABS - UNIDENTIFIED"" will be used instead. " & vbCrLf 
	''				TaskDialog.Show("X","NO")
				Else							
					CurCol  = CurDoc.Create.NewFamilyInstance(CurLine,CurShape,CurLevel,Autodesk.Revit.DB.Structure.StructuralType.Column)
				End If
				
	
				NewList5.Add(i)
				Newlist5.Add(curcol.Id)
				
				ColIdMatch.Add(Newlist5)
			Else 
'				TaskDialog.Show("R", "Found a match")
				ColWarningError = ColWarningError & "     - There was an error preventing the creation of the column specified at CSV line " & i+4 & "." & vbCrLf
			End if

		next
		
	End Sub
	
	Public Sub GetSlabs(curdoc As Document)
		
		
	' Get all floor types loaded in Revit, all levels and assign template types
    Dim collectors As New clsCollectors
    Dim FloorTypeList As List(Of FloorType) = collectors.GetAllFloorType(curdoc)
    Dim TemplateFloorType As FloorType = FloorTypeList(0)
	Dim LevelList As List(Of Level) = collectors.GetAllLevels(curdoc)
	Dim lvl As Level = LevelList(0)
	
	
	'Get all the floors used in the analytical model
	
	Dim FloorName As New List(Of String)	
	
	For i=0 To CSVData.Rows(1).Item(50).ToString-1					
		If CSVData.Rows(3+i).Item(47).ToString = "Floor" Then
			If FloorName.Contains(CSVData.Rows(3+i).Item(49).ToString) Then
			Else FloorName.add(CSVData.Rows(3+i).Item(49).ToString)
			End If
		End If	
	Next	
	
	
	For Each CurFloorName As String In FloorName
		
		'Create the floor type associated with curfloorname		
		
		Dim CurFloorThickness As Double = 0
		Dim FloorTypeExist As Boolean = False
		
		For i=0 To CSVData.Rows(1).Item(29).ToString-1
			If CSVData.Rows(3+i).Item(29).ToString = CurfloorName Then
				CurFloorThickness = CSVData.Rows(3+i).Item(32).ToString 'Get the thickness of the walls from CSV file
			End If					 
		Next
		
		Dim CurFloorTypeName As String = CSVData.Rows(0).Item(0).ToString & " - " & CurFloorThickness.ToString & """ (" & CurFloorName & ")"
		
		'Deleted prohibited character from string
		If CurFloorTypeName.Contains(":") Or CurFloorTypeName.Contains(";") Or CurFloorTypeName.Contains("<") Or CurFloorTypeName.Contains(">") Or CurFloorTypeName.Contains("[") Or CurFloorTypeName.Contains("]") Then
			CurFloorTypeName = CurFloorTypeName.Replace(":","")
			CurFloorTypeName = CurFloorTypeName.Replace(";","")
			CurFloorTypeName = CurFloorTypeName.Replace("<","")
			CurFloorTypeName = CurFloorTypeName.Replace(">","")
			CurFloorTypeName = CurFloorTypeName.Replace("[","")
			CurFloorTypeName = CurFloorTypeName.Replace("]","")			
		End If
		
		
		Dim CurFloorType As FloorType = Nothing
		
		'Check if the floor type is already existing
		For Each FT As FloorType In FloorTypeList
			If FT.Name = CurFloorTypeName Then
				FloorTypeExist = True
				CurFloorType = Ft
			End If
		Next
				
		'If it doesn't exist, create it
		If FloorTypeExist = False Then
			CurFloorType = TemplateFloorType.Duplicate(CurFloorTypeName)			
			Dim CurLayer As CompoundStructure = CurFloorType.GetCompoundStructure
			CurLayer.SetLayerWidth(0,CurFloorThickness/12)
			CurFloorType.SetCompoundStructure(CurLayer)
		End If		
	
		For i=0 To CSVData.Rows(1).Item(50).ToString-1
			If CSVData.Rows(3+i).Item(47).ToString = "Floor" Then
				If CSVData.Rows(3+i).Item(49).ToString = CurFloorName Then 
					
					Dim SlabOutline As New CurveArray
					Dim SlabZ As Double = 9999
					
					For u=0 To CSVData.Rows(3+i).Item(48).ToString-1 
						If CSVData.Rows(3+i+u).Item(53)< SlabZ Then
							SlabZ = CSVData.Rows(3+i+u).Item(53)
						End If
					Next
					
					For u=0 To CSVData.Rows(3+i).Item(48).ToString-1 
						Dim X1 As Double = CSVData.Rows(3+i+u).Item(51)
						Dim Y1 As Double = CSVData.Rows(3+i+u).Item(52)
						Dim Z1 As Double = CSVData.Rows(3+i+u).Item(53)
						Dim X2 As New Double
						Dim Y2 As New Double
						Dim Z2 As New Double
						Dim Point1 As New XYZ(X1,Y1,SlabZ)
						
						If u <> CSVData.Rows(3+i).Item(48).ToString-1 Then
							 X2 = CSVData.Rows(3+i+u+1).Item(51)
							 Y2 = CSVData.Rows(3+i+u+1).Item(52)
							 Z2 = CSVData.Rows(3+i+u+1).Item(53)
	
						Else X2 = CSVData.Rows(3+i).Item(51)
							 Y2 = CSVData.Rows(3+i).Item(52)
							 Z2 = CSVData.Rows(3+i).Item(53)
						End If
						
						Dim Point2 As New XYZ(X2,Y2,SlabZ)
						
						SlabOutline.Append(Line.CreateBound(Point1,Point2))				
						
					Next
										
'					Dim VerticalVector As XYZ = XYZ.BasisZ
					Dim Offset As Double = 9999
					
					For Each Curlevel As Level In LevelList
						If Math.Abs(CSVData.Rows(3+i).Item(53).ToString - Curlevel.Elevation) < Offset Then
							Offset = Math.Abs(CSVData.Rows(3+i).Item(53).ToString - Curlevel.Elevation)
							lvl = Curlevel
						End If	
					Next
					
					Try
						curdoc.Create.NewFloor(SlabOutline,CurFloorType,lvl,True)
'						curdoc.Create.NewSlab(SlabOutline,lvl,Slopex,0.1,True)
					Catch ex as autodesk.revit.Exceptions.InvalidOperationException
							TaskDialog.Show("?","Come see me if you encounter that error. Thanks :)")
						Throw
					End Try


				End If
				
			End If
		Next
		
	next
		
	End Sub
	
	Public Sub GetBeams(curdoc As Document)
				
				
		Dim BeamList As New List(Of list(Of String))  'List of X1,Y1,Z1,X2,Y2,Z2,shape
		Dim NewList As New List(Of String)
		Dim ShapeList As New List(Of String)
		
		Dim VersionPath As String = ""
		
		'Check for Revit version		
		If Application.Application.VersionNumber = "2015" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2015\"
		ElseIf Application.Application.VersionNumber = "2016" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2016\"
		ElseIf Application.Application.VersionNumber = "2017" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2017\"
		ElseIf Application.Application.VersionNumber = "2018" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2018\"
		ElseIf Application.Application.VersionNumber = "2019" Then
			VersionPath = "C:\ProgramData\Autodesk\RVT 2019\"
		End If
		
		'Get all already loaded family and their names in separate lists
		Dim collector As New clsCollectors
		Dim LoadedBeamList As List(Of Family) = collector.getAllBeamFamilies(curdoc)
		Dim LoadedBeamListName As New List(Of String)
		
		For Each Fam As Family In LoadedBeamList
			LoadedBeamListName.Add(Fam.Name)
		Next
		
		'Make sure that at least one round and one rectangular column is loaded, and assign templates
		Dim RectangularTemplate As FamilySymbol = Nothing
		Dim AllRectangular As New List(Of FamilySymbol)
		Dim tempList As New HashSet(Of ElementId) 'List(Of ElementId)
		
		If LoadedBeamListName.Contains("Concrete-Rectangular Beam") Then
			For Each fam As Family In LoadedBeamList
				If fam.Name = "Concrete-Rectangular Beam" Then
					
					tempList = fam.GetFamilySymbolIds
					
					For Each elid As ElementId In tempList
						AllRectangular.Add(DirectCast(CurDoc.GetElement(elid),FamilySymbol))
					Next
					
					RectangularTemplate = DirectCast(CurDoc.GetElement(fam.GetFamilySymbolIds(0)),FamilySymbol)					
				End If
			Next
		Else curdoc.LoadFamilySymbol(VersionPath & "Libraries\US Imperial\Structural Framing\Concrete\Concrete-Rectangular Beam.rfa","12 x 24",New FamilyLoadOptions(),RectangularTemplate)
		End If


		
		'Store the col CSV file data in memory, and create a separate list containing only the shapes		
		For i=0 To CSVData.Rows(1).Item(35).ToString-1
			If CSVData.Rows(3+i).Item(42) >=45 And CSVData.Rows(3+i).Item(43).ToString <> "" Then
				NewList = New List(Of String)
				NewList.Add(CSVData.Rows(3+i).Item(36).ToString)   '
				NewList.Add(CSVData.Rows(3+i).Item(37).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(38).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(39).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(40).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(41).ToString)
				NewList.Add(CSVData.Rows(3+i).Item(43).ToString)
				ShapeList.Add(CSVData.Rows(3+i).Item(43).ToString)
				BeamList.Add(NewList)
			End if
		Next
	
		
		'Eliminate duplicates from the ShapeList		
		ShapeList = ShapeList.Distinct().ToList
		
		'get dimensions of shapes that are not steel		
		Dim ShapeDims As New List(Of List(Of String))
		
		For i=0 To CSVData.Rows(1).Item(20).ToString-1
			If ShapeList.Contains(CSVData.Rows(3+i).Item(20).ToString) Then
				NewList = New List(Of String)
				NewList.add(CSVData.Rows(3+i).Item(20).ToString)
				NewList.add(CSVData.Rows(3+i).Item(22).ToString)
				NewList.add(CSVData.Rows(3+i).Item(23).ToString)
				ShapeDims.Add(NewList)
			End If
		Next
				
		
		'For each shape used, check what type of shape it is
		' list of possible types: I-shape / Channel / T / Angle / DblAngle / Box / Pipe / Rectangular / Circle / DblChannel / Other 
		
		Dim TypeList As New List(Of String)
		
		For Each CurShape As String In ShapeList
			For i=0 To CSVData.Rows(1).Item(20).ToString-1
				If CSVData.Rows(3+i).Item(20).ToString = CurShape Then
					TypeList.Add(CSVData.Rows(3+i).Item(21).ToString)
				End If
			Next
		Next
		
		
		'Associate each Shape with its Revit symbol
		
		Dim NameSymbolPairs As New List(Of List(Of Object))
		Dim ShapePath As String = ""
		Dim VarList As New List(Of object)
		Dim CurSymbol As FamilySymbol = Nothing
		Dim b As Double = 0
		Dim h As Double = 0
		Dim Loaded As Boolean = False
		Dim IdentificationCounter As Integer = 1
		
		'Create an template steel shape for poorly named shapes in CSV file\
			Dim BaseShape As FamilySymbol = Nothing
			CurDoc.LoadFamilySymbol(VersionPath & "Libraries\US Imperial\Structural Framing\Steel\W-Wide Flange.rfa","W12X26",New FamilyLoadOptions(), BaseShape)

		For i=0 To TypeList.Count-1
			
			VarList = New  List(Of Object)
			
'			TaskDialog.Show(i, ShapeList(i) &"/"& TypeList(i))

			'If it's a steel shape
			If TypeList(i) = "I-Shape" Or TypeList(i) = "Channel" Or TypeList(i)= "Angle" Or TypeList(i)= "DblAngle" Or TypeList(i)= "Box" Or TypeList(i)= "Pipe" Or TypeList(i)= "DblChannel" Then				
				
				If TypeList(i)= "I-Shape" Then
					ShapePath = "Libraries\US Imperial\Structural Framing\Steel\W-Wide Flange.rfa"
				ElseIf TypeList(i)= "Channel" Then
					ShapePath = "Libraries\US Imperial\Structural Framing\Steel\C-Channel.rfa"
				ElseIf TypeList(i)= "Angle" Then
					ShapePath = "Libraries\US Imperial\Structural Framing\Steel\L-Angle.rfa"
				ElseIf TypeList(i)= "DblAngle" Then
					ShapePath = "Libraries\US Imperial\Structural Framing\Steel\LL-Double Angle.rfa"
				ElseIf TypeList(i)= "Box" Then
					ShapePath = "Libraries\US Imperial\Structural Framing\Steel\HSS-Hollow Structural Section.rfa"
				ElseIf TypeList(i)= "Pipe" Then
					ShapePath = "Libraries\US Imperial\Structural Framing\Steel\HSS-Round Structural Tubing.rfa"
				ElseIf TypeList(i)= "DblChannel" Then
					ShapePath = "Libraries\US Imperial\Structural Framing\Steel\Double C-Channel.rfa"
				End If
				
				Dim CurName As String =""
				
				If ShapeList(i).EndsWith("-1") Then
					CurName = ShapeList(i).Substring(0,ShapeList(i).Length-2)
				Else
					CurName = ShapeList(i)
				End If
				
				Loaded = CurDoc.LoadFamilySymbol(VersionPath & ShapePath,CurName, CurSymbol)
				
				
				If Loaded = False Then
					CurSymbol = BaseShape.Duplicate(CSVData.Rows(0).Item(0).ToString & " - UNIDENTIFIED " & IdentificationCounter)
					BeamWarningError = BeamWarningError & "     - """ & shapeList(i) & """ is not a valid Revit shape. """ & CurSymbol.Name & """ will be used instead. " & vbCrLf 
					IdentificationCounter = IdentificationCounter+1
				End If
				
				CurSymbol.Activate()
				
				VarList.Add(ShapeList(i))
				VarList.Add(CurSymbol)
				
				NameSymbolPairs.Add(VarList)
				
				
			'If it's a concrete or similar material, then	
			ElseIf TypeList(i)= "Rectangular" Or TypeList(i)= "Other" 	Then
				
				Dim BeamName As String = CSVData.Rows(0).Item(0).ToString & " - " & ShapeList(i)
				
				'Deleted prohibited character from string
				If BeamName.Contains(":") Or BeamName.Contains(";") Or BeamName.Contains("<") Or BeamName.Contains(">") Or BeamName.Contains("[") Or BeamName.Contains("]") Then
					BeamName = BeamName.Replace(":","")
					BeamName = BeamName.Replace(";","")
					BeamName = BeamName.Replace("<","")
					BeamName = BeamName.Replace(">","")
					BeamName = BeamName.Replace("[","")
					BeamName = BeamName.Replace("]","")			
				End If
				
				
				Dim FamSymExist As Boolean = False
					
					For Each FamSym As FamilySymbol In AllRectangular
						If FamSym.Name = BeamName Then
							FamSymExist = True
							VarList.Add(ShapeList(i))
							VarList.Add(famSym)
				
							NameSymbolPairs.Add(VarList)
						End If
					Next
					
					If FamSymExist = False Then
						CurSymbol= RectangularTemplate.Duplicate(BeamName)
						
						For u=0 To ShapeDims.Count-1
							If ShapeDims(u)(0) = ShapeList(i) Then
								b = ShapeDims(u)(1)/12
								h = ShapeDims(u)(2)/12
							End If
						Next
						
						CurSymbol.Activate()
						
						CurSymbol.GetOrderedParameters(1).Set(h)
						CurSymbol.GetOrderedParameters(2).Set(b)
							
						VarList.Add(ShapeList(i))
						VarList.Add(CurSymbol)
				
						NameSymbolPairs.Add(VarList)
								
					End If		

				
			End If

		Next

	
	Dim LevelList As List(Of Level) = collector.GetAllLevels(CurDoc)

		
		'Place all the beams
		For i=0 To BeamList.Count-1

			Dim CurStart As New XYZ(BeamList(i)(0),BeamList(i)(1),BeamList(i)(2))
			Dim CurEnd As New XYZ(BeamList(i)(3),BeamList(i)(4),BeamList(i)(5))
			Dim CurLine As Curve = Line.CreateBound(CurStart,CurEnd)

			Dim Curshape As FamilySymbol = Nothing

			For u=0 To NameSymbolPairs.Count-1
				If NameSymbolPairs(u)(0) = BeamList(i)(6) Then
					CurShape=NameSymbolPairs(u)(1)
				End If
			Next
				
			Dim curlevel As Level = LevelList(0)
			Dim LevelOffset As Double = 999
			Dim CurLowPoint As Double = Math.Min(CDbl(BeamList(i)(2)),CDbl(BeamList(i)(5)))
				
			For Each lvl As Level In LevelList
				If Math.Abs(lvl.Elevation-CurLowPoint)<LevelOffset Then
					curlevel = lvl
					LevelOffset = Math.Abs(lvl.Elevation-CurLowPoint)
				End If
			Next
			
			If Curshape is Nothing then
				TaskDialog.Show(i, BeamList(i)(6))
			End If
			
			Dim CurBeam As FamilyInstance = CurDoc.Create.NewFamilyInstance(CurLine,CurShape,CurLevel,Autodesk.Revit.DB.Structure.StructuralType.Beam)

		next
					
	End Sub	
	
	Public Sub GetWalls(curdoc As Document)
		
	'Get all wall types loaded in Revit	
	Dim collectors As New clsCollectors
	Dim WallTypeList As List(Of WallType) = collectors.GetAllWallTypes(curdoc)
    Dim TemplatewallType As wallType = WallTypeList(0)
    	
	Dim WallName As New List(Of String)	
	
	'Collect all the types of wall used in the analytical model
	For i=0 To CSVData.Rows(1).Item(50).ToString-1					
		If CSVData.Rows(3+i).Item(47).ToString = "Wall" Then
			If WallName.Contains(CSVData.Rows(3+i).Item(49).ToString) Then
			Else WallName.add(CSVData.Rows(3+i).Item(49).ToString)
			End If
		End If	
	Next	
		
	Dim CurX As New List(Of String)
	Dim CurY As New List(Of String)
	Dim CurZ As New List(Of String)	
	
	'Collect all the shell points from the analytical model
	For i=0 To CSVData.Rows(1).Item(50).ToString-1	
		CurX.Add(CSVData.Rows(3+i).Item(51).ToString)
		CurY.Add(CSVData.Rows(3+i).Item(52).ToString)
		CurZ.Add(CSVData.Rows(3+i).Item(53).ToString)
	Next
	
	Dim TempCurveArray As New CurveArray
	Dim curLine As Line = Nothing
	Dim ShellLines As New List(Of Curve)
	Dim LineList As New List(Of Curve)
	
	Dim ExistingPlane As Boolean = False
	Dim CurPlaneDValue As New Double
	Dim PlaneValuesList As New List(Of List(Of Object)) ' list of (Origin first, normal second, d value third (ax+by+cz =d)), one per plane
	Dim PlaneValues As New List(Of Double)
	Dim PlaneList As New List(Of List(Of Curve))
	Dim PrintOut As String =""
	
		
	For Each CurWallName As String In WallName
			
		'Create the wall type necessary in Revit
		
		Dim CurWallThickness As Double = 0
		Dim WallTypeExist As Boolean = False
		
		For i=0 To CSVData.Rows(1).Item(29).ToString-1
			If CSVData.Rows(3+i).Item(29).ToString = CurWallName Then
				CurWallThickness = CSVData.Rows(3+i).Item(32).ToString 'Get the thickness of the walls from CSV file
			End If					 
		Next
		
		Dim CurWallTypeName As String = CSVData.Rows(0).Item(0).ToString & " - " & CurWallThickness.ToString & """ (" & CurWallName & ")"
		
		'Deleted prohibited character from string
		If CurWallTypeName.Contains(":") Or CurWallTypeName.Contains(";") Or CurWallTypeName.Contains("<") Or CurWallTypeName.Contains(">") Or CurWallTypeName.Contains("[") Or CurWallTypeName.Contains("]") Then
			CurWallTypeName = CurWallTypeName.Replace(":","")
			CurWallTypeName = CurWallTypeName.Replace(";","")
			CurWallTypeName = CurWallTypeName.Replace("<","")
			CurWallTypeName = CurWallTypeName.Replace(">","")
			CurWallTypeName = CurWallTypeName.Replace("[","")
			CurWallTypeName = CurWallTypeName.Replace("]","")			
		End If
		
		Dim CurWallType As WallType = Nothing
		
		'Check if the wall type is already existing
		For Each WT As WallType In WallTypeList
			If WT.Name = CurWallTypeName Then
				WallTypeExist = True
				CurWallType = wt
			End If
		Next
				
		'If it doesn't exist, create it
		If WallTypeExist = False Then
			CurWallType = TemplatewallType.Duplicate(CurWallTypeName)			
			Dim CurLayer As CompoundStructure = CurWallType.GetCompoundStructure
			CurLayer.SetLayerWidth(0,CurWallThickness/12)
			CurWallType.SetCompoundStructure(CurLayer)
		End If		
			
		PlaneValuesList = New List(Of List(Of Object))
		PlaneValues = New List(Of Double)
		PlaneList = New List(Of List(Of Curve))
		
		Dim Counter As Integer = 0
			
		' -------- Getting all the relevant lines (of type "wall" and name "curwallName" in a list named LineList	------

		
		For i=0 To CSVData.Rows(1).Item(50).ToString-1          'Count of shell elements in model
			
			
			If CSVData.Rows(3+i).Item(47).ToString = "Wall" Then    'only proceed if the shell is of type "wall"
				If CSVData.Rows(3+i).Item(49).ToString = CurWallName Then    'only proceed if the wall is of name "CurwallName"
					
					Counter=Counter+1
					
					'Empty the data from the previous shell
					ShellLines = New List(Of Curve)
					TempCurveArray = New CurveArray
					
					For u = 0 To CSVData.Rows(3+i).Item(48).ToString-1 'for each points of the shell (4 usually for a wall, potentially more for slabs
						If u = CSVData.Rows(3+i).Item(48).ToString-1 Then  
							'if it's the last point of the shell, then create a line between it and the first point to close the loop
							
							Dim PtA As New XYZ(CurX(i+u),CurY(i+u),CurZ(i+u))
							Dim PtB As New XYZ(CurX(i),CurY(i),CurZ(i))
				
							CurLine = Line.CreateBound(PtA,PtB)
					
							ShellLines.Add(CurLine)
							
						Else
							For a = u+1 To CSVData.Rows(3+i).Item(48).ToString-1 
								If (CurX(i+u) = curX(i+a) And curY(i+u)=CurY(i+a)) Or CurZ(i+u)=CurZ(i+a) Then 
								' If the two points have the same x and y, or the same z, then the line between the two of them is an edge of the shell (not a diagonal, which would have at most x OR y the same)	
								
									Dim PtA As New XYZ(CurX(i+u),CurY(i+u),CurZ(i+u))
									Dim PtB As New XYZ(CurX(i+a),CurY(i+a),CurZ(i+a))					
								
									CurLine = Line.CreateBound(PtA,PtB)		
								
									ShellLines.Add(CurLine)
	
									a = CSVData.Rows(3+i).Item(48).ToString-1
								End If
							Next							
						
						End if
					Next
			
					For each Curlines As curve in ShellLines
						TempCurveArray.Append(curLines)
					Next
					
					Dim PlanePt1 As XYZ = ShellLines(0).GetEndPoint(0)
					Dim PlanePt2 As XYZ = ShellLines(0).GetEndPoint(1)
					Dim PlanePt3 As XYZ = ShellLines(1).GetEndPoint(1)
			
					'Sort the newly created lines into lists of coplanar lines
			
			'Create the plane containing the current shell			

					Dim Curplane As Plane = Plane.CreateByThreePoints(PlanePt1,PlanePt2,PlanePt3)
'					Dim CurPlane As Plane = curdoc.application.Create.NewPlane(TempCurveArray)
						
					CurPlaneDValue = CurPlane.Origin.X*CurPlane.Normal.X+CurPlane.Origin.y*CurPlane.Normal.y+CurPlane.Origin.z*CurPlane.Normal.z
			
					ExistingPlane = False
			
					'Check the d value as well as the normal vector against the ones already existing in the PlaneValuesList. If there is a match, then it is not a new plane but an already defined one
					If PlaneValuesList.Count>0 Then
						For a=0 To PlaneValuesList.Count-1
							If (PlaneValuesList(a)(2) = CurPlaneDValue Or PlaneValuesList(a)(2) = -CurPlaneDValue) And _
								((PlaneValuesList(a)(1).x = CurPlane.Normal.X And PlaneValuesList(a)(1).Y = CurPlane.Normal.Y And PlaneValuesList(a)(1).Z = CurPlane.Normal.Z)  Or _
								(PlaneValuesList(a)(1).x = -CurPlane.Normal.X And PlaneValuesList(a)(1).Y = -CurPlane.Normal.Y And PlaneValuesList(a)(1).Z = -CurPlane.Normal.Z)) then
								ExistingPlane = True
						
								'Add the current shell lines to the matching plane
						
								For Each CurShellLines As Curve In ShellLines
									PlaneList(a).Add(CurShellLines)
								Next
						
							End If
						Next
					End If
			
					If ExistingPlane = False Then
						'Since there was no match in the PlaneValuesList, it means ShellLines is the first shell of its plane, so add it to PlaneValuesList and PlaneList
						Dim TempPlaneValues As New List(Of Object)
				
						TempPlaneValues.Add(CurPlane.origin)
						TempPlaneValues.Add(CurPlane.normal)
						TempPlaneValues.Add(CurPlaneDValue)
						PlaneValuesList.Add(TempPlaneValues)
				
						PlaneList.Add(ShellLines)
							
					End If
			
				End If
			End If	
		Next

		
		' ------ All the lines should be retrieved and sorted into planes -----
		
		' ------ Let's filter those lines and delete overlapping/duplicate ones ------
		
		
		Dim TempPt1 As New XYZ(-1000,1000,-1000)
		Dim TempPT2 As New xyz(1000,-1000,1000)
		Dim Replacementline As Line = Line.CreateBound(TempPt1,TempPT2) 'this will be used to locate which lines should be deleted forever
		
		Dim Pt1 As New XYZ(0,0,0)
		Dim Pt2 As New XYZ(0,0,0)
		Dim Pt3 As New XYZ(0,0,0)
		Dim Pt4 As New XYZ(0,0,0)
		
		
		For a=0 To PlaneList.Count-1
			
			For i=0 To PlaneList(a).Count-1
				
				Pt1 = PlaneList(a)(i).GetEndPoint(0)
				Pt2 = PlaneList(a)(i).GetEndPoint(1)
				
				For u = 0 To PlaneList(a).Count-1
					
					If u<>i Then
						
						Pt3 = PlaneList(a)(u).GetEndPoint(0)
						Pt4 = PlaneList(a)(u).GetEndPoint(1)	
						
						If (Pt3.X =Pt1.x And Pt3.y =Pt1.y And Pt3.Z =Pt1.z And Pt4.X =Pt2.x And Pt4.y =Pt2.y And Pt4.z =Pt2.z) Or _
							(Pt3.X =Pt2.x And Pt3.y =Pt2.y And Pt3.z =Pt2.Z And Pt4.X =Pt1.x And Pt4.y =Pt1.y And Pt4.z =Pt1.z) Then
							'If the two lines are exact matches, mark the two of them for deletion
							PlaneList(a)(u) = Replacementline
							PlaneList(a)(i) = Replacementline
					
						ElseIf (Pt3.X =Pt1.x And Pt3.y =Pt1.y And Pt3.Z =Pt1.z) or (Pt4.X =Pt2.x And Pt4.y =Pt2.y And Pt4.z =Pt2.z) Or _
								(Pt3.X =Pt2.x And Pt3.y =Pt2.y And Pt3.z =Pt2.Z) Or (Pt4.X =Pt1.x And Pt4.y =Pt1.y And Pt4.z =Pt1.z) Then
							
							'If they have one end point in common but not the other ones, then delete one and change the other one to the substraction of the two
					
								If (Pt3.X =Pt1.x And Pt3.y =Pt1.y And Pt3.Z =Pt1.z) And (Is_Between(Pt1,Pt4,Pt2) Or Is_Between(Pt3,Pt2,Pt4)) Then 
									PlaneList(a)(i)= Line.CreateBound(Pt4,Pt2)
									PlaneList(a)(u)= Replacementline
									Pt1 = PlaneList(a)(i).GetEndPoint(0)
									Pt2 = PlaneList(a)(i).GetEndPoint(1)
									u=0
									
								ElseIf (Pt4.X =Pt1.x And Pt4.y =Pt1.y And Pt4.z =Pt1.z) And (Is_Between(Pt1,Pt3,Pt2) Or Is_Between(Pt1,Pt2,Pt3)) Then 
									PlaneList(a)(u)= Replacementline
									PlaneList(a)(i)= Line.CreateBound(Pt3,Pt2)
									Pt1 = PlaneList(a)(i).GetEndPoint(0)
									Pt2 = PlaneList(a)(i).GetEndPoint(1)
									u=0
									
								ElseIf (Pt2.X =Pt3.x And Pt2.y =Pt3.y And Pt2.z =Pt3.z) And (Is_Between(Pt3,Pt4,Pt1) Or Is_Between(Pt3,Pt1,Pt4)) Then 
									PlaneList(a)(u) = Replacementline
									PlaneList(a)(i) = Line.CreateBound(Pt1,Pt4)
									Pt1 = PlaneList(a)(i).GetEndPoint(0)
									Pt2 = PlaneList(a)(i).GetEndPoint(1)
									u=0
									
								ElseIf (Pt4.X =Pt2.x And Pt4.y =Pt2.y And Pt4.z =Pt2.z) And (Is_Between(Pt2,Pt3,Pt1) Or Is_Between(Pt2,Pt1,Pt3)) Then
									PlaneList(a)(u) = Replacementline
									PlaneList(a)(i) = Line.CreateBound(Pt1,Pt3)
									Pt1 = PlaneList(a)(i).GetEndPoint(0)
									Pt2 = PlaneList(a)(i).GetEndPoint(1)
									u=0
								End if
						Else 
							'IF they don't have any end points in common, check if they still partially overlap
							
							If Is_Between(Pt1,Pt3,Pt2)=True Then
								If Is_Between(Pt1,Pt4,Pt3)=True Or Is_Between(Pt4,Pt1,Pt3)=True Then
									PlaneList(a)(u) = Line.CreateBound(Pt1,Pt4)
									PlaneList(a)(i)=Line.CreateBound(Pt3,Pt2)
									Pt1 = PlaneList(a)(i).GetEndPoint(0)
									Pt2 = PlaneList(a)(i).GetEndPoint(1)
'									u=0

								Elseif Is_Between(Pt3,Pt2,Pt4)=True Or Is_Between(Pt3,Pt4,Pt2)=True Then
									PlaneList(a)(u)=Line.CreateBound(Pt1,Pt3)
									PlaneList(a)(i) =Line.CreateBound(Pt4,Pt2)
									Pt1 = PlaneList(a)(i).GetEndPoint(0)
									Pt2 = PlaneList(a)(i).GetEndPoint(1)
'									u=0
								End If
								
							ElseIf (Is_Between(Pt1,Pt4,Pt2) And Is_Between(Pt4,Pt2,Pt3)) Or (Is_Between(Pt4,Pt1,Pt2) And Is_Between(Pt1,Pt2,Pt3))Then
								PlaneList(a)(u)=Line.CreateBound(Pt1,Pt4)
								PlaneList(a)(i)=Line.CreateBound(Pt3,Pt2)
								Pt1 = PlaneList(a)(i).GetEndPoint(0)
								Pt2 = PlaneList(a)(i).GetEndPoint(1)
'								u=0
								
							ElseIf (Is_Between(Pt3,Pt1,Pt4) And Is_Between(Pt1,Pt4,Pt2)) Or (Is_Between(Pt3,Pt1,Pt2) And Is_Between(Pt1,Pt2,Pt4))Then
								PlaneList(a)(u)=Line.CreateBound(Pt1,Pt3)
								PlaneList(a)(i)=Line.CreateBound(Pt4,Pt2)
								Pt1 = PlaneList(a)(i).GetEndPoint(0)
								Pt2 = PlaneList(a)(i).GetEndPoint(1)
'								u=0

							End If
					
						End if
					End if
				Next
			Next
			
		Next
		
						
		For a=0 To PlaneList.Count-1	
			'Delete all the lines marked for deletion				
			While PlaneList(a).Contains(Replacementline)
				PlaneList(a).Remove(Replacementline)
			End While
			
		Next

		
		' ------ Now that we have all the lines filtered, they need to be rearranged to form closed loops ----- GOOD UP TO HERE 
		
		Dim OrderedPlaneList As New List(Of List(Of List(Of Curve)))
		Dim FollowingLine As Boolean = False
		Dim CurCurve As Curve = Nothing
		Dim LineToBeRemoved As Line = Nothing
		Dim NewList As List(Of Curve) 
		Dim NewListList As List(Of List(Of Curve))
		

		OrderedPlaneList = New List(Of List(Of List(Of Curve)))
		NewList = New List(Of Curve) 
		NewListList = new List(Of List(Of Curve))

		For a= 0 To PlaneList.Count-1 '0 To 0'
			
			NewListList = New List(Of List(Of Curve))
			NewList = New List(Of Curve)
			FollowingLine = false
			
			While  PlaneList(a).Count>0
	
				If FollowingLine = False Then
											
					If NewList IsNot Nothing AndAlso NewList.Count>0 Then
						NewListList.Add(NewList)
					End If
						
					NewList = New List(Of curve)
					curCurve = PlaneList(a)(0)
					NewList.Add(Curcurve)
					PlaneList(a).Remove(curCurve)
						
				Else curcurve = NewList(NewList.Count-1)
				End If	

				FollowingLine = False

				For u=0 To PlaneList(a).Count-1	

					If PlaneList(a)(u).GetEndPoint(0).X=curCurve.GetEndPoint(1).X And PlaneList(a)(u).GetEndPoint(0).y=curCurve.GetEndPoint(1).y And PlaneList(a)(u).GetEndPoint(0).z=curCurve.GetEndPoint(1).z Then				
						NewList.Add(PlaneList(a)(u))
						FollowingLine=True
						LineToBeRemoved = PlaneList(a)(u)
						u=0
						Exit For
							
					ElseIf PlaneList(a)(u).GetEndPoint(1).X=curCurve.GetEndPoint(1).X And PlaneList(a)(u).GetEndPoint(1).y=curCurve.GetEndPoint(1).y And PlaneList(a)(u).GetEndPoint(1).z=curCurve.GetEndPoint(1).z Then
						NewList.Add(Line.CreateBound(PlaneList(a)(u).GetEndPoint(1),PlaneList(a)(u).GetEndPoint(0)))
						FollowingLine=True
						LineToBeRemoved = PlaneList(a)(u)
						u=0
						Exit for
					End If
				Next
	
				If FollowingLine= True Then
					PlaneList(a).Remove(LineToBeRemoved)
				End If
		
			End While
			
			NewListList.Add(Newlist)		
			OrderedPlaneList.Add(NewListList)	
					
		Next

		
		
		' ------ Now that we have all the closed loops sorted out, let's figure out which ones are walls and which ones are doors/windows -----
		
		'First, let's get the min and max points of each closed curves, so that we can then see which curves are inside which other curve
		
		Dim LoopBoundaries As List(Of List(Of List(Of Double)))  'MinX, MinY, MinZ, MaxX, maxY, MaxZ
		Dim newList2 As List(Of Double)
		Dim NewListList2 As List(Of List(Of Double))
		
		
		LoopBoundaries = New List(Of List(Of List(Of Double))) ' idk why I need that line, but it's not working without
		
		For a= 0 To PlaneList.Count-1
			
			NewListList2 = new List(Of List(Of Double))
	
			For u=0 To OrderedPlaneList(a).Count-1
				
				newList2 = New List(Of Double)
				
				For o=0 To 2
					NewList2.Add(99999)
				Next
				For o=0 To 2
					NewList2.Add(-99999)
				Next
					
				For i=0 To OrderedPlaneList(a)(u).Count-1
						
					If OrderedPlaneList(a)(u)(i).GetEndPoint(0).X < NewList2(0) Or OrderedPlaneList(a)(u)(i).GetEndPoint(1).X < NewList2(0) Then
						newList2(0) = Math.Min(OrderedPlaneList(a)(u)(i).GetEndPoint(0).X,OrderedPlaneList(a)(u)(i).GetEndPoint(1).X)
					End If
					
					If OrderedPlaneList(a)(u)(i).GetEndPoint(0).y < NewList2(1) Or OrderedPlaneList(a)(u)(i).GetEndPoint(1).y < NewList2(1) Then
						newList2(1) = Math.Min(OrderedPlaneList(a)(u)(i).GetEndPoint(0).y,OrderedPlaneList(a)(u)(i).GetEndPoint(1).y)
					End If
					
					If OrderedPlaneList(a)(u)(i).GetEndPoint(0).z < NewList2(2) Or OrderedPlaneList(a)(u)(i).GetEndPoint(1).z < NewList2(2) Then
						newList2(2) = Math.Min(OrderedPlaneList(a)(u)(i).GetEndPoint(0).z,OrderedPlaneList(a)(u)(i).GetEndPoint(1).z)
					End If
					
					If OrderedPlaneList(a)(u)(i).GetEndPoint(0).X > NewList2(3) Or OrderedPlaneList(a)(u)(i).GetEndPoint(1).X > NewList2(3) Then
						newList2(3) = Math.Max(OrderedPlaneList(a)(u)(i).GetEndPoint(0).X,OrderedPlaneList(a)(u)(i).GetEndPoint(1).X)
					End If
					
					If OrderedPlaneList(a)(u)(i).GetEndPoint(0).y > NewList2(4) Or OrderedPlaneList(a)(u)(i).GetEndPoint(1).y > NewList2(4) Then
						newList2(4) = Math.Max(OrderedPlaneList(a)(u)(i).GetEndPoint(0).y,OrderedPlaneList(a)(u)(i).GetEndPoint(1).y)
					End If
					
					If OrderedPlaneList(a)(u)(i).GetEndPoint(0).z > NewList2(5) Or OrderedPlaneList(a)(u)(i).GetEndPoint(1).z > NewList2(5) Then
						newList2(5) = Math.Max(OrderedPlaneList(a)(u)(i).GetEndPoint(0).z,OrderedPlaneList(a)(u)(i).GetEndPoint(1).z)
					End If
					
				Next	
				
				NewListList2.Add(newList2)
	
			Next

			LoopBoundaries.Add(NewListList2)
			
		Next
		
		
		'Now, we can compare those min and max 
		
		Dim WallOrOpng As List(Of List(Of Integer)) 'for each plane, we will have a list of integer (one per closed curves). It will be -1 for walls, and the wall closed curve for opng on said wall
		Dim WallBool As Boolean = true
		Dim Newlist3 As List(Of Integer)
		
		WallOrOpng = New List(Of List(Of Integer))
		
		For a= 0 To PlaneList.Count-1
						
			Newlist3= New List(Of Integer)
						
			For u=0 To OrderedPlaneList(a).Count-1
				
				WallBool=True ' the closed curve is a wall until proven otherwise
				
				For i=0 To OrderedPlaneList(a).Count-1
					
					If (LoopBoundaries(a)(u)(0) > LoopBoundaries(a)(i)(0) Or LoopBoundaries(a)(u)(1) > LoopBoundaries(a)(i)(1)) And LoopBoundaries(a)(u)(2) > LoopBoundaries(a)(i)(2) And _
					    (LoopBoundaries(a)(u)(3) < LoopBoundaries(a)(i)(3) or LoopBoundaries(a)(u)(4) < LoopBoundaries(a)(i)(4)) And LoopBoundaries(a)(u)(5) < LoopBoundaries(a)(i)(5) Then
						' if loop(u) is strictly inside loop(i) i
						Newlist3.Add(i)
						WallBool = False   'proven not a wall
					End If
				
				Next
				
				If WallBool = True Then
					Newlist3.Add(-1)
				End If
				
			Next
			
			WallOrOpng.Add(Newlist3)
			
		Next
'		
		
		' ------ Now that the data is sorted and we figured out which closed loop is what, we can finally create the walls and opngs ------ '
		
		Dim CurWall As wall = nothing


		For a= 0 To PlaneList.Count-1
			
'			Dim ca As New CurveArray
'			
'			For i=0 To OrderedPlaneList(a)(0).Count-1
'				ca.Append(OrderedPlaneList(a)(0)(i))
'			Next
'			
'			Dim tmpPlane As Plane = curdoc.Application.Create.NewPlane(ca)
'			Dim tmpSketchPlane As SketchPlane = SketchPlane.Create(curdoc,tmpPlane)
'			
'			For i=0 To OrderedPlaneList(a).Count-1
'				For u=0 To OrderedPlaneList(a)(i).Count-1
'					CurDoc.Create.NewModelCurve(OrderedPlaneList(a)(i)(u),tmpSketchPlane)			
'				Next
'	
'			Next
			
			For u=0 To WallOrOpng(a).Count-1
				
				If WallOrOpng(a)(u) = -1 Then
					
					Try
						CurWall = Wall.Create(curdoc, OrderedPlaneList(a)(u),False)
						curwall.WallType = CurWallType
						
						For i=0 To WallOrOpng(a).Count-1
							If WallOrOpng(a)(i)=u Then
								
								Try
									Pt1 = New XYZ(LoopBoundaries(a)(i)(0),LoopBoundaries(a)(i)(1),LoopBoundaries(a)(i)(2))
									Pt2 = New XYZ(LoopBoundaries(a)(i)(3),LoopBoundaries(a)(i)(4),LoopBoundaries(a)(i)(5))
									curdoc.Create.NewOpening(CurWall,pt1,pt2)
								Catch 
									Pt1 = New XYZ(LoopBoundaries(a)(i)(0),LoopBoundaries(a)(i)(4),LoopBoundaries(a)(i)(2))
									Pt2 = New XYZ(LoopBoundaries(a)(i)(3),LoopBoundaries(a)(i)(1),LoopBoundaries(a)(i)(5))
									curdoc.Create.NewOpening(CurWall,pt1,pt2)									
								End Try
'								curdoc.Create.NewOpening(CurWall,OrderedPlaneList(a)(i)(0).GetEndPoint(0),OrderedPlaneList(a)(i)(1).GetEndPoint(1))
							End If
						Next
						
						
					Catch
						
					End Try
									
				End If
			Next
		
			
		next		
		
		
	Next

		
		
	End Sub
			
	Function Distance(a As XYZ, b As XYZ) As Double
		
		Return Math.Sqrt((a.X-b.X)^2+(a.y-b.y)^2+(a.z-b.z)^2)
		
	End Function
	
	Function Is_Between(a As XYZ, c As XYZ, b As XYZ) As Boolean
		
		If Math.Abs(Distance(a,c)+Distance(c,b)-Distance(a,b))<0.1 And ((c.X<>a.X Or c.Y<>a.Y Or c.Z<>a.Z) OR (c.X<>b.X Or c.Y<>b.Y Or c.Z<>b.Z))  Then
			Return True
			Else return false
		End If
		
	End Function
	
	Function Is_Inside_Loop(MinX1 As Double, MinY1 As Double, MinZ1 As Double, MinX2 As Double, MinY2 As Double, MinZ2 As Double, MaxX1 As Double, MaxY1 As Double, MaxZ1 As Double, MaxX2 As Double, MaxY2 As Double, MaxZ2 As Double) As Boolean
		
		If MinX1>MinX2 And MinY1>MinY2 And MinZ1>MinZ2 And MaxX1<MaxX2 And MaxY1<MaxY2 And MaxZ1<MaxZ2 Then
			Return True
		Else 
			Return False
		End If
		
	End Function
	
	
End Class

Public Class FamilyLoadOptions
        Implements IFamilyLoadOptions

    Public Function OnFamilyFound(ByVal familyInUse As Boolean, ByRef overwriteParameterValues As Boolean) As Boolean Implements IFamilyLoadOptions.OnFamilyFound

		overwriteParameterValues = True
		Return True

    End Function

    Public Function OnSharedFamilyFound(ByVal sharedFamily As Autodesk.Revit.DB.Family, ByVal familyInUse As Boolean, ByRef source As FamilySource, ByRef overwriteParameterValues As Boolean) As Boolean Implements IFamilyLoadOptions.OnSharedFamilyFound

		source = FamilySource.Family
		overwriteParameterValues = True
		Return True

    End Function

End Class

Public Class FailureHandler
	Implements IFailuresPreprocessor
	
	Public ErrorMsg As String = ""
	Public ErrorSeverity As String =""
	Public property ProblemColId As new List(Of ElementId)
	
	Public Function PreprocessFailures(failuresAccessor As FailuresAccessor) As FailureProcessingResult  Implements IFailuresPreprocessor.PreprocessFailures
		
		Dim FailureMessages As IList(Of FailureMessageAccessor) = failuresAccessor.GetFailureMessages()
		Dim FatalError As Boolean = false
		
'		TaskDialog.Show("MSG", FailureMessages.Count)
		
		For Each Msg As FailureMessageAccessor In FailureMessages
			'Delete the warning messages that can be ignored, and the others will be rolled back
			
			Dim FailureId As FailureDefinitionId = Msg.GetFailureDefinitionId
			Dim FailedElementId As List(Of ElementId) = Msg.GetFailingElementIds
			
			
			Try 
				ErrorMsg = msg.GetDescriptionText
			Catch 
				ErrorMsg = "Unknown Error"
			End try
			
			Try 
				Dim FailureSeverity As FailureSeverity = Msg.GetSeverity
				ErrorSeverity = FailureSeverity.ToString
				
				If FailureSeverity = FailureSeverity.Warning Then
					failuresAccessor.DeleteWarning(Msg)
				Else 
					
				If FailureSeverity = FailureSeverity.Error then		
					ProblemColId.AddRange(FailedElementId)	
					FatalError=True
				End If
				
				End If
			Catch
				
			End Try
			
		Next
		
'		For Each elid As ElementId In ProblemColId
'			TaskDialog.Show("x",elid.ToString)
'		Next
							
		If FatalError=True Then
			Return FailureProcessingResult.ProceedWithRollBack
		else
			Return FailureProcessingResult.Continue
		End If
		
	End Function
	
	
	
End Class
			
