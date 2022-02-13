'
' Created by SharpDevelop.
' User: michael
' Date: 4/2/2015
' Time: 12:21 PM
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

Public Class clsCollectors
	public Function getAllSheets(curDoc As Document) As List(Of ViewSheet)
		'get all sheets
		Dim m_colViews As New FilteredElementCollector(curDoc)
		m_colViews.OfCategory(BuiltInCategory.OST_Sheets)
		
		Dim m_Sheets As New List(Of Viewsheet)
		For Each x As Viewsheet In m_colViews.ToElements
			m_Sheets.Add(x)
		Next
		
		Return m_Sheets
	End Function
	
	public Function getAllSections(curDoc As Document) As List(Of View)
		'get all viewports
		Dim vpCollector As New FilteredElementCollector(curDoc)
		vpCollector.OfCategory(BuiltInCategory.OST_Views)
		
		
		'output structural views to list
		Dim SectionList As New List(Of View)
		For Each curView As View In vpCollector
			If curView.ViewType = ViewType.Section Then
				If curView.Name <> "Structural Section" AndAlso curView.Name <> "Architectural Section" AndAlso curView.Name <> "Site Section" _
					AndAlso curView.Name <>"Structural Framing Elevation" AndAlso curView.Name <>"Architectural Elevation" _
					AndAlso curView.Name <>"Structural Building Elevation" AndAlso curView.Name <>"3/4"" steel" Then
					SectionList.Add(curView)
				End if
			End If		
	
		Next
		
		Return SectionList
		
	End Function
	
	
	Public Function GetAllFamilies(curdoc As Document) As List(Of Family)
        'Get all loaded families in the project
        Dim m_colFamily As New FilteredElementCollector(curdoc)
        m_colFamily.OfClass(GetType(Family))

        Dim m_Family As New List(Of Family)
        For Each F As Family In m_colFamily
            m_Family.Add(F)
        Next

        Return m_Family
	End Function
	
	
	Public Function getAllColumnFamilies(CurDoc As Document) As List(Of family)
		'get all column families loaded
		Dim m_colcolFamily As New FilteredElementCollector(CurDoc) 

		m_colcolFamily.OfClass(GetType(Family))'.OfCategory(BuiltInCategory.OST_StructuralColumns)
		
		Dim ColumnList As New List(Of Family)
		
		For Each CurFam As family In m_colcolFamily
			If CurFam.FamilyCategory.Name = "Structural Columns" then
				ColumnList.Add(CurFam)
			End if
		Next
		
		Return ColumnList
		
	End Function
	
	Public Function getAllBeamFamilies(CurDoc As Document) As List(Of family)
		'get all column families loaded
		Dim m_colBeamFamily As New FilteredElementCollector(CurDoc) 

		m_colBeamFamily.OfClass(GetType(Family))'.OfCategory(BuiltInCategory.OST_StructuralColumns)
		
		Dim BeamList As New List(Of Family)
		
		For Each CurFam As family In m_colBeamFamily
			If CurFam.FamilyCategory.Name = "Structural Framing" then
				BeamList.Add(CurFam)
			End if
		Next
		
		Return BeamList
		
	End Function
	
	Public function GetAllLevels(curdoc As Document) As List(Of Level)
		
		Dim m_colLevel As New FilteredElementCollector(curdoc)
		m_colLevel.OfClass(GetType(Level))
		
		Dim LevelList As New List(Of Level)
		For Each lvl As Level In m_colLevel
			LevelList.Add(lvl)			
		Next
		
		Return LevelList
		
	End Function
	
	Public Function GetAllGrids(curdoc As Document) As List(Of Grid)
		
		Dim m_colGrid As New FilteredElementCollector(curdoc)
		m_colGrid.OfClass(GetType(Grid))
		
		Dim GridList As New List(Of Grid)
		For Each Gr As Grid In m_colGrid
			GridList.Add(Gr)
		Next
		
		Return GridList
		
	End Function
		
	Public Function GetAllWallTypes(curdoc As Document) As List(Of WallType)
			
		Dim m_WallType As New FilteredElementCollector(curdoc)
		m_WallType.Ofclass(GetType(WallType))
		
		Dim WallTypeList As New List(Of WallType)
		For Each WT As WallType In m_WallType
			WallTypeList.Add(WT)
		Next
		
		Return WallTypeList
			
	End Function
	
	Public Function GetAllFloorType(curdoc As Document) As List( Of FloorType)
		
		Dim m_FloorType As New FilteredElementCollector(curdoc)
		m_FloorType.Ofclass(GetType(FloorType))
		
		Dim FloorTypeList As New List(Of FloorType)
		For Each FT As FloorType In m_FloorType
			FloorTypeList.Add(FT)
		Next
		
		Return FloorTypeList
		
	End Function
	
End Class
