Attribute VB_Name = "mAutoComplete"
Option Explicit

Public Enum eAutoCompleteBehavior
    acbListOnly
    acbFile
    acbFolder
    acbMultiSelect = 4
End Enum

Public Sub ac_Change(poComboText As Object, ByVal piLastKeyDown As Integer, Optional ByVal Behavior As eAutoCompleteBehavior)
    Dim i As Long, liLen As Long, lsText As String
    On Error Resume Next
    Dim loCBox As ComboBox

    Set loCBox = poComboText
    If piLastKeyDown = 220 And Behavior Mod acbMultiSelect > acbListOnly Then Exit Sub
    Select Case piLastKeyDown
        Case vbKeyDelete, vbKeyBack
        Case Else
            With poComboText
                Select Case Behavior Mod acbMultiSelect
                    Case acbFolder, acbFile
                        lsText = .Text
                        liLen = InStrRev(lsText, ";") + 1
                        If BitIsSet(Behavior, acbMultiSelect) Then lsText = Mid$(lsText, liLen)
                        If StrComp(Right$(lsText, 1), "\") <> 0 Then lsText = FindFirst(lsText, Behavior Mod acbMultiSelect = acbFolder)

                        If Len(lsText) > 0 Then
                            If BitIsSet(Behavior, acbMultiSelect) Then lsText = Mid$(.Text, 1, liLen - 1) & lsText
                        End If
                End Select
                
                liLen = Len(lsText)
                
                If liLen = 0 Then
                    If Not loCBox Is Nothing Then
                        With loCBox
                            lsText = .Text
                            liLen = Len(lsText)
                            For i = 0 To .ListCount - 1
                                If StrComp(Mid$(.List(i), 1, liLen), lsText, vbTextCompare) = 0 Then
                                    If .List(i) = "Clear This List" Then Exit Sub
                                    lsText = .List(i)
                                    Exit For
                                End If
                            Next
                            If i >= .ListCount Then lsText = ""
                        End With
                    End If
                End If
                
                If Len(lsText) > 0 Then
                    liLen = .SelStart
                    .Text = lsText
                    .SelStart = liLen
                    .SelLength = Len(.Text) - liLen
                End If
            End With
    End Select
End Sub

Public Sub ac_KeyPress(ByVal poComboText As Object, KeyAscii As Integer, Optional ByVal Behavior As eAutoCompleteBehavior)
    On Error Resume Next
    Dim liTemp As Long
    With poComboText
        Select Case Behavior Mod acbMultiSelect
            Case acbFile, acbFolder
                If KeyAscii = 92 Then
                    liTemp = InStr(.SelStart + 1, .Text, Chr$(KeyAscii))
                    'Debug.Print .SelStart + .SelLength, liTemp
                    If liTemp > .SelStart + .SelLength + 1 Then liTemp = 0
                    If InStr(liTemp + 1, .Text, Chr$(KeyAscii)) > 0 Then liTemp = 0
                    If liTemp > 0 Then
                        .SelStart = liTemp
                        .SelLength = Len(.Text) - liTemp
                        KeyAscii = 0
                    End If
                End If
        End Select
        If BitIsSet(Behavior, acbMultiSelect) Then
            If KeyAscii = 59 Then
                liTemp = InStr(.SelStart + 1, .Text, Chr$(KeyAscii))
                If liTemp > 0 Then
                    .SelStart = liTemp - 1
                    .SelLength = 1
                Else
                    .SelStart = Len(.Text)
                End If
            End If
        End If
    End With
End Sub
