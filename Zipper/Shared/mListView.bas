Attribute VB_Name = "mListView"
Option Explicit

Private Const SortKeySuffix = "SKey"

Public Sub SortListView(ByVal poListView As ListView, ByVal poColHeader As ColumnHeader, Optional ByVal piOrder As ListSortOrderConstants = -1)
    On Error Resume Next
    Select Case poColHeader.Key
        Case "ReceiveDate", "CompleteDate", "InvoiceDate", "Since", "Amount", "Balance", "Size"
            EnsureSortKey poListView, poColHeader.Key
            Set poColHeader = poListView.ColumnHeaders(poColHeader.Key & SortKeySuffix)
    End Select
    
    If poListView.SortKey = poColHeader.SubItemIndex And piOrder = -1 Then 'And poListView.Sorted = True
        If poListView.SortOrder = lvwAscending Then poListView.SortOrder = lvwDescending Else poListView.SortOrder = lvwAscending
        poListView.Sorted = True
    Else
        If piOrder = -1 Then piOrder = lvwAscending
        poListView.SortOrder = piOrder
        poListView.SortKey = poColHeader.SubItemIndex
        poListView.Sorted = True
    End If
    poListView.Sorted = False
    poListView.SelectedItem.EnsureVisible
End Sub


Public Sub EnsureSortKey(poListView As ListView, psColumnKey As String)
    Dim lsNewKey As String
    Dim liOldIndex  As Integer
    Dim liNewIndex  As Integer
    Dim loCurLI     As ListItem
    Dim lsText As String
    Dim liTemp As Long
    
    lsNewKey = psColumnKey & SortKeySuffix
    liOldIndex = poListView.ColumnHeaders(psColumnKey).SubItemIndex
    On Error Resume Next
    poListView.ColumnHeaders.Add(, lsNewKey, , 0).Tag = SortKeySuffix
    poListView.Sorted = False
    liNewIndex = poListView.ColumnHeaders(lsNewKey).SubItemIndex
    Select Case psColumnKey
        Case "ReceiveDate", "CompleteDate", "Since", "InvoiceDate"
            For Each loCurLI In poListView.ListItems
                loCurLI.SubItems(liNewIndex) = CDbl(CDate(loCurLI.SubItems(liOldIndex)))
            Next
        Case Else
            For Each loCurLI In poListView.ListItems
                lsText = loCurLI.SubItems(liOldIndex)
                If Asc(lsText) = 36 Then lsText = Right$(lsText, Len(lsText) - 1)
                liTemp = InStr(1, lsText, " ")
                If liTemp > 0 Then lsText = Left$(lsText, liTemp)
                loCurLI.SubItems(liNewIndex) = Format(lsText, "0000000.000000")
            Next
    End Select
End Sub

