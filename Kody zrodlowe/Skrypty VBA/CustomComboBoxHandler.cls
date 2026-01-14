VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomComboBoxHandler"
' -------------------------------------------------------------------------------------
' Class: CustomComboBoxHandler
' Description: Enhanced wrapper for Excel Shape-based ComboBoxes.
' Supports Key-Value pairs and automatic data population from files, sheets, or ranges.
' -------------------------------------------------------------------------------------

Option Explicit

Private m_comboBox As Shape
Private m_itemKeys As Variant

' Initialize the handler with a specific ComboBox shape on a worksheet
Public Sub SetControl(ByVal targetSheet As Worksheet, ByVal controlName As String)
    On Error Resume Next
    Set m_comboBox = targetSheet.Shapes(controlName)
    m_itemKeys = Array()
End Sub

' Clear all items from the ComboBox and reset keys
Public Sub Clear()
    If m_comboBox Is Nothing Then Exit Sub
    m_comboBox.ControlFormat.RemoveAllItems
    m_itemKeys = Array()
End Sub

' Add an item with an associated unique key (Variant)
Public Function AddItem(ByVal itemText As String, ByVal itemKey As Variant) As Boolean
    On Error GoTo ErrHandler
    If m_comboBox Is Nothing Then Exit Function

    ' Add text to UI
    m_comboBox.ControlFormat.AddItem itemText

    ' Store key in internal array
    If IsEmpty(m_itemKeys) Or (LBound(m_itemKeys) > UBound(m_itemKeys)) Then
        m_itemKeys = Array(itemKey)
    Else
        Dim newSize As Long
        newSize = UBound(m_itemKeys) + 1
        ReDim Preserve m_itemKeys(newSize)
        m_itemKeys(newSize) = itemKey
    End If

    AddItem = True
    Exit Function
ErrHandler:
    AddItem = False
End Function

' Get the key associated with the currently selected item
Public Property Get SelectedKey() As Variant
    If m_comboBox Is Nothing Then Exit Property
    Dim selectedIdx As Long
    selectedIdx = m_comboBox.ControlFormat.ListIndex

    If selectedIdx > 0 Then
        SelectedKey = m_itemKeys(selectedIdx - 1)
    Else
        SelectedKey = Null
    End If
End Property

' Populate ComboBox with file names from a directory based on a mask
Public Function PopulateFromFiles(ByVal folderPath As String, ByVal fileMask As String) As Boolean
    On Error GoTo ErrHandler
    If m_comboBox Is Nothing Then Exit Function

    Me.Clear
    Dim fileName As String
    fileName = Dir(folderPath & fileMask)

    Do While fileName <> ""
        Me.AddItem fileName, folderPath & fileName
        fileName = Dir
    Loop

    PopulateFromFiles = True
    Exit Function
ErrHandler:
    PopulateFromFiles = False
End Function

' Populate ComboBox with all worksheet names in a given workbook
Public Function PopulateFromSheets(ByVal targetWorkbook As Workbook) As Boolean
    On Error GoTo ErrHandler
    If m_comboBox Is Nothing Or targetWorkbook Is Nothing Then Exit Function

    Me.Clear
    Dim ws As Worksheet
    For Each ws In targetWorkbook.Worksheets
        Me.AddItem ws.Name, ws.Name
    Next ws

    PopulateFromSheets = True
    Exit Function
ErrHandler:
    PopulateFromSheets = False
End Function

' Populate ComboBox from a range (e.g., a header row)
Public Function PopulateFromRange(ByVal headerRange As Range) As Boolean
    On Error GoTo ErrHandler
    If m_comboBox Is Nothing Or headerRange Is Nothing Then Exit Function

    Me.Clear
    Dim cellOffset As Long: cellOffset = 0

    ' Iterate horizontally until an empty cell is found
    Do While Not IsEmpty(headerRange.Offset(0, cellOffset).Value)
        Dim val As String
        val = CStr(headerRange.Offset(0, cellOffset).Value)
        Me.AddItem val, cellOffset
        cellOffset = cellOffset + 1
    Loop

    PopulateFromRange = True
    Exit Function
ErrHandler:
    PopulateFromRange = False
End Function