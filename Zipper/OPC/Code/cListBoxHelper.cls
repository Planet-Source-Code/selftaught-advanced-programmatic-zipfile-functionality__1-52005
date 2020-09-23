VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListBoxAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_GETFONT As Long = &H31
Private Const WM_SETREDRAW = &HB
Private Const ERR_INVALIDPROC = 5&

'Private Const WM_USER = &H400
Private Const DIR_NORMALFILES = &H0
'Private Const DIR_READONLY = &H8001
'Private Const DIR_HIDDEN = &H8002
'Private Const DIR_SYSTEM = &H8004
Private Const DIR_DIRECTORIES = &H8010
'Private Const DIR_ARCHIVED = &H8020
Private Const DIR_DRIVES = &HC000



'Private Const LB_ADDFILE As Long = &H196&
Private Const LB_ADDSTRING As Long = &H180&
Private Const LB_DELETESTRING As Long = &H182&
Private Const LB_DIR As Long = &H18D&
'Private Const LB_DST_ADDR_USE_DSTADDR_FLAG As Long = &H8&
'Private Const LB_DST_ADDR_USE_SRCADDR_FLAG As Long = &H4&
'Private Const LB_DST_MASK_LATE_FLAG As Long = &H20&
Private Const LB_ERR As Long = -1&
'Private Const LB_ERRSPACE As Long = -2&
Private Const LB_FINDSTRING As Long = &H18F&
Private Const LB_FINDSTRINGEXACT As Long = &H1A2&
'Private Const LB_GETANCHORINDEX As Long = &H19D&
'Private Const LB_GETCARETINDEX As Long = &H19F&
Private Const LB_GETCOUNT As Long = &H18B&
Private Const LB_GETCURSEL As Long = &H188&
'Private Const LB_GETHORIZONTALEXTENT As Long = &H193&
Private Const LB_GETITEMDATA As Long = &H199&
'Private Const LB_GETITEMHEIGHT As Long = &H1A1&
'Private Const LB_GETITEMRECT As Long = &H198&
'Private Const LB_GETLOCALE As Long = &H1A6&
Private Const LB_GETSEL As Long = &H187&
Private Const LB_GETSELCOUNT As Long = &H190&
Private Const LB_GETSELITEMS As Long = &H191&
Private Const LB_GETTEXT As Long = &H189&
Private Const LB_GETTEXTLEN As Long = &H18A&
Private Const LB_GETTOPINDEX As Long = &H18E&
Private Const LB_INITSTORAGE As Long = &H1A8&
Private Const LB_INSERTSTRING As Long = &H181&
'Private Const LB_ITEMFROMPOINT As Long = &H1A9&
'Private Const LB_MSGMAX As Long = &H1A8&
'Private Const LB_MULTIPLEADDSTRING As Long = &H1B1&
'Private Const LB_OKAY As Long = 0&
Private Const LB_RESETCONTENT As Long = &H184&
'Private Const LB_SELECTSTRING As Long = &H18C&
'Private Const LB_SELITEMRANGE As Long = &H19B&
'Private Const LB_SELITEMRANGEEX As Long = &H183&
'Private Const LB_SETANCHORINDEX As Long = &H19C&
'Private Const LB_SETCARETINDEX As Long = &H19E&
'Private Const LB_SETCOLUMNWIDTH As Long = &H195&
'Private Const LB_SETCOUNT As Long = &H1A7&
Private Const LB_SETCURSEL As Long = &H186&
Private Const LB_SETHORIZONTALEXTENT As Long = &H194&
Private Const LB_SETITEMDATA As Long = &H19A&
'Private Const LB_SETITEMHEIGHT As Long = &H1A0&
'Private Const LB_SETLOCALE As Long = &H1A5&
Private Const LB_SETSEL As Long = &H185&
'Private Const LB_SETTABSTOPS As Long = &H192&
Private Const LB_SETTOPINDEX As Long = &H197&
'Private Const LB_SRC_ADDR_USE_DSTADDR_FLAG As Long = &H2&
'Private Const LB_SRC_ADDR_USE_SRCADDR_FLAG As Long = &H1&
'Private Const LB_SRC_MASK_LATE_FLAG As Long = &H10&


Private Const ITEM_MARGIN = 4
Private Const CHECKBOX_WIDTH = 14

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
                            ByVal Hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any _
                           ) As Long

Private Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" ( _
                            ByVal Hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long _
                         ) As Long
    Private Const WM_VSCROLL = &H115
    Private Const SB_BOTTOM = 7

Private Declare Function GetDC Lib "User32" ( _
                            ByVal Hwnd As Long _
                          ) As Long
Private Declare Function ReleaseDC Lib "User32" ( _
                            ByVal Hwnd As Long, _
                            ByVal hDC As Long _
                          ) As Long

Private Declare Function DrawText Lib "User32" Alias "DrawTextA" ( _
                            ByVal hDC As Long, _
                            ByVal lpStr As String, _
                            ByVal nCount As Long, _
                            lpRect As RECT, _
                            ByVal wFormat As Long _
                          ) As Long
    Private Const DT_SINGLELINE = &H20
    Private Const DT_CALCRECT = &H400

    Private Type RECT
       Left As Long
       Top As Long
       Right As Long
       Bottom As Long
    End Type

Private Declare Function SelectObject Lib "gdi32" ( _
                            ByVal hDC As Long, _
                            ByVal hObject As Long _
                          ) As Long

Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" ( _
                            ByVal Hwnd As Long, _
                            ByVal nIndex As Long _
                          ) As Long
    Private Const GWL_STYLE = (-16)
    Private Const WS_VSCROLL = &H200000

Private Declare Function GetSystemMetrics Lib "User32" ( _
                            ByVal nIndex As Long _
                           ) As Long
    Private Const SM_CXVSCROLL = 2

' --- private class variables ---

Private mbIsCheckBox As Boolean
Private miMaxWidthPixels As Long
Private miHFont As Long
Private miHwnd As Long
Private mbSingleSel As Boolean
Private mbRedrawLocked As Boolean

' Call this sub before you use other methods of this class
Public Sub Init(ByVal LBox As ListBox)
    With LBox
        miHwnd = .Hwnd
        mbIsCheckBox = .Style = vbListBoxCheckbox
        mbSingleSel = .MultiSelect = vbMultiSelectNone
        'Get the handle to be used in api calls such as
        'calculating the width of text
        miHFont = SendMessage(miHwnd, WM_GETFONT, 0&, 0&)
    End With
End Sub

' The following routine adds a string to the list box
' and displays the horizontal scroll bar if required
Public Function AddItem(ByRef Item As String, Optional ByVal Index As Long = -1, Optional ByVal OnlyIfUnique As Boolean) As Long
    On Error Resume Next
    AddItem = LB_ERR
    If OnlyIfUnique Then
        If FindString(Item, True) > LB_ERR Then Exit Function
    End If
    
    If Index > -1& Then
        AddItem = SendMessage(miHwnd, LB_INSERTSTRING, Index, ByVal Item)
    Else
        AddItem = SendMessage(miHwnd, LB_ADDSTRING, 0&, ByVal Item)
    End If
   
    If AddItem <= LB_ERR Then Exit Function
   
    Dim liWidth As Long
    ' --- calculating the width of the currently added item ---
    liWidth = GetItemWidth(Item)
   
    ' --- determining whether we need to display the horizontal scroll bar ---
    If liWidth > miMaxWidthPixels Then
        miMaxWidthPixels = liWidth
        SetHorExtent
    End If
End Function

Public Sub Clear()
    SendMessageLong miHwnd, LB_RESETCONTENT, 0&, 0&
    miMaxWidthPixels = 0&
    SendMessageLong miHwnd, LB_SETHORIZONTALEXTENT, 0&, 0&
End Sub

' Use this property to change the caption of a list box item
Public Property Let List(ByVal Index As Long, Item As String)
    Dim liOldWidth As Long
    Dim liNewWidth As Long
    
    liOldWidth = GetItemWidth(Me.List(Index))

    SendMessageLong miHwnd, WM_SETREDRAW, -1&, 0&
    If SendMessage(miHwnd, LB_INSERTSTRING, Index, ByVal Item) <= LB_ERR Then Exit Property
    SendMessageLong miHwnd, LB_DELETESTRING, Index + 1, 0&
    LockRedraw mbRedrawLocked
    
    liNewWidth = GetItemWidth(Item)
    
    If liOldWidth >= miMaxWidthPixels Then
        If liNewWidth >= liOldWidth Then
            miMaxWidthPixels = liNewWidth
            SetHorExtent
        Else
            RefreshHScroll
        End If
    Else
        If liNewWidth > miMaxWidthPixels Then
            miMaxWidthPixels = liNewWidth
            SetHorExtent
        End If
    End If
End Property

Private Sub SetHorExtent()
    SendMessageLong miHwnd, LB_SETHORIZONTALEXTENT, miMaxWidthPixels + ITEM_MARGIN + IIf(mbIsCheckBox, CHECKBOX_WIDTH, 0), 0
End Sub

Public Property Get ListCount() As Long
    ListCount = SendMessageLong(miHwnd, LB_GETCOUNT, 0&, 0&)
End Property

Public Sub RemoveItem(ByVal Index As Integer)
    Dim liWidth As Long
    liWidth = GetItemWidth(Me.List(Index))
    If SendMessageLong(miHwnd, LB_DELETESTRING, Index, 0&) <= LB_ERR Then Exit Sub
   
    If Me.ListCount = 0 Then
        miMaxWidthPixels = 0&
        SendMessageLong miHwnd, LB_SETHORIZONTALEXTENT, 0&, 0&
    Else
        If liWidth >= miMaxWidthPixels Then RefreshHScroll
        ' we remove the item with the maximum width
        ' and must recalculate the maximum width and the
        ' horizontal extent of our list box
    End If
End Sub

' recalculates the maximum width for all items and
' displays the horizontal scroll bar if required
Public Sub RefreshHScroll()
    Dim ltRECT As RECT
    Dim hOldFont As Long
    Dim hDC As Long
    Dim I As Long
    Dim liWidth As Long
    
    ' We use the same technique like in the flGetItemWidth function,
    ' but sligtly optimize it: we don't need to select/deselect the
    ' required font for each item - we can do it at once outside the cycle
    hDC = GetDC(miHwnd)
    hOldFont = SelectObject(hDC, miHFont) ' selecting the required font
    
    miMaxWidthPixels = 0
    With ltRECT
        For I = 0 To ListCount - 1
            DrawText hDC, Me.List(I), -1, ltRECT, DT_SINGLELINE + DT_CALCRECT
            liWidth = .Right - .Left
            If liWidth > miMaxWidthPixels Then miMaxWidthPixels = liWidth
        Next
    End With
    ' restoring the previous font
    Call SelectObject(hDC, hOldFont)
    ReleaseDC miHwnd, hDC
    SetHorExtent
End Sub

Private Function GetItemWidth(Item As String) As Long
   Dim hDC As Long
   Dim ltRECT As RECT
   Dim hOldFont As Long
   If LenB(Item) = 0 Then Exit Function
   hDC = GetDC(miHwnd) ' retrieving HDC for the listbox
   hOldFont = SelectObject(hDC, miHFont) ' selecting the required font
   ' if you specify the DT_CALCRECT flag,
   ' DrawText only Determines the width and height of the rectangle
   ' required to display the text:
   DrawText hDC, Item, -1, ltRECT, DT_SINGLELINE + DT_CALCRECT
   GetItemWidth = ltRECT.Right - ltRECT.Left
   ' restoring the previous state
   Call SelectObject(hDC, hOldFont)
   ReleaseDC miHwnd, hDC
End Function

' Returns the minimum width of the list box
' when the horizontal scroll bar disappears
' NOTE: the return value in pixels
Public Property Get MinWidthNoHScroll() As Long
    Dim lbHasVScroll As Boolean
    lbHasVScroll = GetWindowLong(miHwnd, GWL_STYLE) And WS_VSCROLL
    If mbIsCheckBox Then MinWidthNoHScroll = CHECKBOX_WIDTH
    If lbHasVScroll Then
        MinWidthNoHScroll = MinWidthNoHScroll + GetSystemMetrics(SM_CXVSCROLL)
    Else
        MinWidthNoHScroll = MinWidthNoHScroll + miMaxWidthPixels
    End If
End Property

Public Function RemoveSelection() As Boolean
    'This function only works with a simple MultiSelect listbox, not extended
    Dim liIndices() As Long
    Dim I As Long
    On Error GoTo errout
    
    GetSelIndices liIndices
    For I = UBound(liIndices) To 0 Step -1
        Me.RemoveItem liIndices(I)
    Next
    RemoveSelection = True
    RefreshHScroll
errout:
End Function

Public Property Get SelIndex() As Long
    SelIndex = SendMessageLong(miHwnd, LB_GETCURSEL, 0&, 0&)
End Property

Public Property Let SelIndex(ByVal NewIndex As Long)
    SendMessageLong miHwnd, LB_SETCURSEL, NewIndex, 0&
End Property

Public Function GetSelIndices(Indices() As Long) As Boolean
    'This property only works with a simple MultiSelect listbox, not extended
    Dim ItemIndexes() As Long
    Dim liNumItems As Long
    
    If miHwnd = 0 Then Exit Function
    
    If mbSingleSel Then
        liNumItems = SelIndex
        If liNumItems > -1 Then
            ReDim ItemIndexes(0 To 0)
            ItemIndexes(0) = liNumItems
        End If
    Else
        liNumItems = Me.SelCount
        If liNumItems <= 0 Then Exit Function Else GetSelIndices = True
        If liNumItems Then
            ReDim ItemIndexes(liNumItems - 1)
            SendMessage miHwnd, LB_GETSELITEMS, liNumItems, ItemIndexes(0)
        End If
    End If
    Indices = ItemIndexes
End Function

Public Sub SetSelIndices(Indices() As Long)
    On Error Resume Next
    Dim I As Long
    Selected(-1) = False
    For I = LBound(Indices) To UBound(Indices)
        Selected(Indices(I)) = True
    Next
End Sub

Public Property Get SelCount() As Long
    SelCount = SendMessageLong(miHwnd, LB_GETSELCOUNT, 0&, 0&)
End Property

Public Property Get SelItems(Items() As String) As Boolean
    'This property only works with a simple MultiSelect listbox, not extended
    Dim liIndices() As Long
    Dim lsReturn() As String
    Dim I As Long
    On Error GoTo errout
    
    SelItems = GetSelIndices(liIndices)
    If Not SelItems Then Exit Property
    
    ReDim lsReturn(0 To UBound(liIndices))
    For I = 0 To UBound(liIndices)
        lsReturn(I) = Me.List(liIndices(I))
    Next
    Items = lsReturn
errout:
End Property

Public Property Get List(ByVal Index As Long) As String
    Dim liLen As Long
    liLen = SendMessageLong(miHwnd, LB_GETTEXTLEN, Index, 0&)
    If liLen <= LB_ERR Then Err.Raise ERR_INVALIDPROC
    List = Space$(liLen)
    List = Left$(List, SendMessage(miHwnd, LB_GETTEXT, Index, ByVal List))
End Property

Public Property Get FindString(Item As String, ByVal Exact As Boolean) As Long
    FindString = SendMessage(miHwnd, IIf(Exact, LB_FINDSTRINGEXACT, LB_FINDSTRING), -1, ByVal Item)
End Property

'Call this sub to speed up inserts of many items
'If you overestimate, the extra memory is allocated.
'If you underestimate, normal allocation is used for the additional items
Public Sub InitStorage(ByVal NumItems As Long, ByVal AvgLength As Long)
    SendMessageLong miHwnd, LB_INITSTORAGE, NumItems, AvgLength * 2
End Sub

Public Sub LockRedraw(Optional ByVal Locked As Boolean)
    SendMessageLong miHwnd, WM_SETREDRAW, CLng(Not Locked), 0&
    mbRedrawLocked = Locked
End Sub

Public Property Get Selected(ByVal Index As Long) As Boolean
    Selected = SendMessage(miHwnd, LB_GETSEL, Index, 0&) > 0&
End Property

Public Property Let Selected(ByVal Index As Long, ByVal NewVal As Boolean)
    If Not mbSingleSel Then
        SendMessageLong miHwnd, LB_SETSEL, Abs(NewVal), Index
    Else
        If NewVal Then SelIndex = Index
    End If
End Property

Public Property Get TopIndex() As Long
    TopIndex = SendMessageLong(miHwnd, LB_GETTOPINDEX, 0&, 0&)
End Property

Public Property Let TopIndex(ByVal NewVal As Long)
    SendMessageLong miHwnd, LB_SETTOPINDEX, NewVal, 0&
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
    ItemData = SendMessageLong(miHwnd, LB_GETITEMDATA, Index, 0&)
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal NewVal As Long)
    SendMessageLong miHwnd, LB_SETITEMDATA, Index, NewVal
End Property

Public Sub ListFiles(FileSpec As String, Optional ByVal IncludeDrives As Boolean = True, Optional ByVal IncludeSubDirectories As Boolean = True)
    Clear
    If IncludeDrives Then SendMessage miHwnd, LB_DIR, DIR_DRIVES, ByVal FileSpec
    If IncludeSubDirectories Then SendMessage miHwnd, LB_DIR, DIR_DIRECTORIES, ByVal FileSpec
    SendMessage miHwnd, LB_DIR, DIR_NORMALFILES, ByVal FileSpec
End Sub
