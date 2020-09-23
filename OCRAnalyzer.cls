VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OCRAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim Grid(1 To 5, 1 To 5)
    Dim StartY As Long
    Dim EndY As Long
    Dim StartX As Long
    Dim EndX As Long
    Dim SectionInterval As Long
    Dim XSectionInterval As Long
Private Property Let Status(ByVal NewValue As String)
    Let frmOCR.lblStatus.Caption = NewValue
    frmOCR.Refresh
End Property
Private Function Progress() As ProgressBar
    Set Progress = frmOCR.ProgressBar1
End Function
Private Function MaxAllowance() As Long
    Let MaxAllowance = 0 - Val(frmOCR.txtToleranceLevel.Text)
End Function
Private Function MaxPosAllowance() As Long
    Let MaxPosAllowance = Val(frmOCR.txtToleranceLevel.Text)
End Function
Private Function MaxBadSectors() As Long
    Let MaxBadSectors = Val(frmOCR.txtMaxBadSectors.Text)
End Function
Private Function InDebugMode() As Boolean
    Let InDebugMode = frmOCR.InDebugMode
End Function
    '189x173
Private Sub FillColor(XSector As Long, YSector As Long, NumberOfPixels As Long, pbPictureBox As PictureBox)
    If Not InDebugMode Then Exit Sub
    Dim X As Long
    Dim Y As Long
    With pbPictureBox
        For Y = ((XSector * 38) - 38) To (XSector * 38)
            For X = ((YSector * 35) - 35) To (YSector * 35)
                pbPictureBox.PSet (X, Y), QBColorForPixelValue(NumberOfPixels)
            Next X
        Next Y
    End With
End Sub
Private Function QBColorForPixelValue(PixelValue As Long) As Long
    If PixelValue = 0 Then
        Let QBColorForPixelValue = vbWhite
    Else
        Let QBColorForPixelValue = (PixelValue * 3) + 76000
    End If
End Function
Private Sub DrawDebugGrids()
    Dim X As Long '37.8
    Dim Y As Long '34.6
    With frmOCR
        For X = 0 To .pbDebug1.ScaleWidth - 1 Step 37.8
            For Y = 0 To .pbDebug1.ScaleHeight - 1 Step 1
                .pbDebug1.PSet (X, Y), QBColor(10)
            Next Y
        Next X
        For X = 0 To .pbDebug2.ScaleWidth - 1 Step 37.8
            For Y = 0 To .pbDebug2.ScaleHeight - 1 Step 1
                .pbDebug2.PSet (X, Y), QBColor(10)
            Next Y
        Next X
        For Y = 0 To .pbDebug1.ScaleHeight - 1 Step 34.6
            For X = 0 To .pbDebug1.ScaleWidth - 1 Step 1
                .pbDebug1.PSet (X, Y), QBColor(10)
            Next X
        Next Y
        For Y = 0 To .pbDebug2.ScaleHeight - 1 Step 34.6
            For X = 0 To .pbDebug2.ScaleWidth - 1 Step 1
                .pbDebug2.PSet (X, Y), QBColor(10)
            Next X
        Next Y
    End With
End Sub
Private Function DrawingOK() As Boolean
    Let DrawingOK = True
    If StartY = 0 Or _
        EndY = 0 Or _
        StartX = 0 Or _
        EndX = 0 Then
            MsgBox "Something isn't right?", vbInformation, "Oops"
            Let DrawingOK = False
    End If
End Function
Private Sub ResetGrid()
    For Y = 1 To 5
        For X = 1 To 5
            Let Grid(X, Y) = 0
        Next X
    Next Y
End Sub
Public Function Analyze(ThePic As PictureBox, Optional Training As Boolean = False) As String
    Dim tmrAt As Long
    Let tmrAt = Timer
    Let frmOCR.OCRImage.AutoRedraw = False
    If Training Then
        If DrawingOK Then
            SaveLetter UCase(InputBox("What letter is this?"))
            Let Analyze = "Trained!"
            Exit Function
        End If
    End If
    GatherSizeInformation ThePic
    If Not DrawingOK Then
        Let Analyze = "Ooops"
        Let frmOCR.Label2.Caption = ""
        Let ThePic.AutoRedraw = True
        Exit Function
    End If
    ResetGrid
    ReadPixels ThePic
    DrawGridLinesAndBox ThePic
    DisplayGridValues
    DrawUserDebugInfo
    Let Analyze = LetterForWeight()
    Let frmOCR.OCRImage.AutoRedraw = True
    Debug.Print Timer - tmrAt, "seconds"
End Function
Private Sub DrawUserDebugInfo()
    If Not InDebugMode Then Exit Sub
    Dim X As Long
    Dim Y As Long
    For X = 1 To 5
        For Y = 1 To 5
            FillColor X, Y, CLng(Grid(X, Y)), frmOCR.pbDebug1
        Next Y
    Next X
End Sub
Private Sub ReadPixels(ThePic As PictureBox)
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Let Status = "Scanning Pixels"
    Let Progress.Value = 0
    Let Progress.Max = EndX - StartX
    For X = StartX To EndX - 1 Step CLng(Left(frmOCR.txtSampleRate.Text, 1))
        Let Progress.Value = X - (StartX - 1)
        For Y = StartY To EndY - 1 Step CLng(Right(frmOCR.txtSampleRate.Text, 1))
            DoEvents
            If ThePic.Point(X, Y) = 0 Then
                ThePic.PSet (X, Y), QBColor(12)
                Select Case Y
                    Case StartY To StartY + SectionInterval
                        Select Case X
                            Case StartX To StartX + XSectionInterval
                                Let Grid(1, 1) = Grid(1, 1) + 1
                            Case StartX + XSectionInterval + 1 To StartX + (XSectionInterval * 2)
                                Let Grid(1, 2) = Grid(1, 2) + 1
                            Case StartX + (XSectionInterval * 2) + 1 To StartX + (XSectionInterval * 3)
                                Let Grid(1, 3) = Grid(1, 3) + 1
                            Case StartX + (XSectionInterval * 3) + 1 To StartX + (XSectionInterval * 4)
                                Let Grid(1, 4) = Grid(1, 4) + 1
                            Case StartX + (XSectionInterval * 4) + 1 To StartX + (XSectionInterval * 5)
                                Let Grid(1, 5) = Grid(1, 5) + 1
                        End Select
                    Case StartY + SectionInterval + 1 To StartY + (SectionInterval * 2)
                        Select Case X
                            Case StartX To StartX + XSectionInterval
                                Let Grid(2, 1) = Grid(2, 1) + 1
                            Case StartX + XSectionInterval + 1 To StartX + (XSectionInterval * 2)
                                Let Grid(2, 2) = Grid(2, 2) + 1
                            Case StartX + (XSectionInterval * 2) + 1 To StartX + (XSectionInterval * 3)
                                Let Grid(2, 3) = Grid(2, 3) + 1
                            Case StartX + (XSectionInterval * 3) + 1 To StartX + (XSectionInterval * 4)
                                Let Grid(2, 4) = Grid(2, 4) + 1
                            Case StartX + (XSectionInterval * 4) + 1 To StartX + (XSectionInterval * 5)
                                Let Grid(2, 5) = Grid(2, 5) + 1
                        End Select
                    Case StartY + (SectionInterval * 2) + 1 To StartY + (SectionInterval * 3)
                        Select Case X
                            Case StartX To StartX + XSectionInterval
                                Let Grid(3, 1) = Grid(3, 1) + 1
                            Case StartX + XSectionInterval + 1 To StartX + (XSectionInterval * 2)
                                Let Grid(3, 2) = Grid(3, 2) + 1
                            Case StartX + (XSectionInterval * 2) + 1 To StartX + (XSectionInterval * 3)
                                Let Grid(3, 3) = Grid(3, 3) + 1
                            Case StartX + (XSectionInterval * 3) + 1 To StartX + (XSectionInterval * 4)
                                Let Grid(3, 4) = Grid(3, 4) + 1
                            Case StartX + (XSectionInterval * 4) + 1 To StartX + (XSectionInterval * 5)
                                Let Grid(3, 5) = Grid(3, 5) + 1
                        End Select
                    Case StartY + (SectionInterval * 3) + 1 To StartY + (SectionInterval * 4)
                        Select Case X
                            Case StartX To StartX + XSectionInterval
                                Let Grid(4, 1) = Grid(4, 1) + 1
                            Case StartX + XSectionInterval + 1 To StartX + (XSectionInterval * 2)
                                Let Grid(4, 2) = Grid(4, 2) + 1
                            Case StartX + (XSectionInterval * 2) + 1 To StartX + (XSectionInterval * 3)
                                Let Grid(4, 3) = Grid(4, 3) + 1
                            Case StartX + (XSectionInterval * 3) + 1 To StartX + (XSectionInterval * 4)
                                Let Grid(4, 4) = Grid(4, 4) + 1
                            Case StartX + (XSectionInterval * 4) + 1 To StartX + (XSectionInterval * 5)
                                Let Grid(4, 5) = Grid(4, 5) + 1
                        End Select
                    Case StartY + (SectionInterval * 4) + 1 To StartY + (SectionInterval * 5)
                        Select Case X
                            Case StartX To StartX + XSectionInterval
                                Let Grid(5, 1) = Grid(5, 1) + 1
                            Case StartX + XSectionInterval + 1 To StartX + (XSectionInterval * 2)
                                Let Grid(5, 2) = Grid(5, 2) + 1
                            Case StartX + (XSectionInterval * 2) + 1 To StartX + (XSectionInterval * 3)
                                Let Grid(5, 3) = Grid(5, 3) + 1
                            Case StartX + (XSectionInterval * 3) + 1 To StartX + (XSectionInterval * 4)
                                Let Grid(5, 4) = Grid(5, 4) + 1
                            Case StartX + (XSectionInterval * 4) + 1 To StartX + (XSectionInterval * 5)
                                Let Grid(5, 5) = Grid(5, 5) + 1
                        End Select
                End Select
            End If
        Next Y
    Next X
End Sub
Private Sub DrawGridLinesAndBox(ThePic As PictureBox)
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Let i = 0
    For X = StartY To EndY + SectionInterval Step SectionInterval
        Let i = i + 1
        If i <= 6 Then
            For Y = 0 To ThePic.ScaleWidth - 1
                ThePic.PSet (Y, X), QBColor(10)
            Next Y
        End If
    Next X
    Let i = 0
    For X = StartX To EndX + XSectionInterval Step XSectionInterval
        Let i = i + 1
        If i <= 6 Then
            For Y = 0 To ThePic.ScaleHeight - 1
                ThePic.PSet (X, Y), QBColor(10)
            Next Y
        End If
    Next X
    For X = StartY To EndY
        ThePic.PSet (StartX, X), QBColor(9)
    Next X
    For X = StartY To EndY
        ThePic.PSet (EndX, X), QBColor(9)
    Next X
    For X = StartX To EndX
        ThePic.PSet (X, StartY), QBColor(9)
    Next X
    For X = StartX To EndX
        ThePic.PSet (X, EndY), QBColor(9)
    Next X
End Sub
Private Sub DisplayGridValues()
    Dim X As Long
    Dim Y As Long
    Let frmOCR.Label2.Caption = ""
    For X = 1 To 5
        For Y = 1 To 5
            Let frmOCR.Label2.Caption = frmOCR.Label2.Caption & WithSpaces(4, CStr(Grid(X, Y)))
        Next Y
        Let frmOCR.Label2.Caption = frmOCR.Label2.Caption & vbCrLf
    Next X
End Sub
Public Function WithSpaces(NumOfCharacters As Long, StringToSpace As String) As String
    Let WithSpaces = StringToSpace & Space(NumOfCharacters - Len(StringToSpace))
End Function
Private Function SectorForGrid(X As Long, Y As Long) As Long
    Let SectorForGrid = ((X * 5) - 5) + Y
End Function
Private Function LetterForWeight() As String
    Dim LetterData() As String
    Dim ThisChoicePassed As Boolean
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim NumberOfChoices As Long
    Dim FileGridData() As String
    Dim Difference As Long
    Dim NotChecked As Boolean
    Dim ScaleValue As Double
    Dim PossibleLetters() As LetterChoice
    Dim TheChoice As New LetterChoice
    Dim NumberOfBadSectors As Long
    Dim strTemp As String
    If Dir(App.Path & "\letterinfo.txt") = "" Then
        Exit Function
    End If
    Let Status = "Reading Letter File"
    Open App.Path & "\letterinfo.txt" For Input As #1
        Do Until EOF(1)
            Let i = i + 1
            ReDim Preserve LetterData(i)
            Line Input #1, LetterData(i)
        Loop
    Close #1
    Let Status = "Gathering Matches"
    Let Progress.Max = UBound(LetterData)
    For i = 1 To UBound(LetterData)
        Let Progress.Value = i
        Let FileGridData = Split(LetterData(i), ",")
        Let ThisChoicePassed = True
        Let ScaleValue = (EndY - StartY) / CLng(FileGridData(26))
        Let NumberOfBadSectors = 0
        For X = 1 To 5
            For Y = 1 To 5
                If (Grid(X, Y) - FileGridData(SectorForGrid(X, Y)) > MaxPosAllowance) Or (Grid(X, Y) - FileGridData(SectorForGrid(X, Y)) < MaxAllowance) Then
                    If (Grid(X, Y) - (FileGridData(SectorForGrid(X, Y)) * ScaleValue) > MaxPosAllowance) Or (Grid(X, Y) - (FileGridData(SectorForGrid(X, Y)) * ScaleValue) < MaxAllowance) Then
                        Let NumberOfBadSectors = NumberOfBadSectors + 1
                        If NumberOfBadSectors > MaxBadSectors Then
                            Let ThisChoicePassed = False
                        End If
                    End If
                End If
            Next Y
            If Not ThisChoicePassed Then Exit For
        Next X
        If ThisChoicePassed Then
            For X = 1 To 5
                For Y = 1 To 5
                    FillColor X, Y, CLng(FileGridData(SectorForGrid(X, Y))), frmOCR.pbDebug2
                Next Y
            Next X
            Let Difference = 0
            For X = 1 To 5
                For Y = 1 To 5
                    If Abs(Grid(X, Y) - FileGridData(X * Y)) > MaxPosAllowance Or Abs(Grid(X, Y) - FileGridData(X * Y)) < MaxAllowance Then
                        Let Difference = Difference + Abs(Grid(X, Y) - (FileGridData(X * Y) * ScaleValue))
                    Else
                        Let Difference = Difference + Abs(Grid(X, Y) - FileGridData(X * Y))
                    End If
                Next Y
            Next X
            Let NumberOfChoices = NumberOfChoices + 1
            ReDim Preserve PossibleLetters(NumberOfChoices)
            Set PossibleLetters(NumberOfChoices) = New LetterChoice
            Let PossibleLetters(NumberOfChoices).DifferenceValue = Difference
            Let PossibleLetters(NumberOfChoices).Letter = Left(LetterData(i), 1)
            Let PossibleLetters(NumberOfChoices).FileData = LetterData(i)
        End If
    Next i
    Let Status = "Choosing match"
    If NumberOfChoices > 0 Then
        Let Progress.Value = 0
        Let Progress.Max = NumberOfChoices
    End If
    Let NotChecked = True
    For i = 1 To NumberOfChoices
        Let Progress.Value = i
        If PossibleLetters(i).DifferenceValue < TheChoice.DifferenceValue Or NotChecked Then
            Let NotChecked = False
            Let TheChoice.DifferenceValue = PossibleLetters(i).DifferenceValue
            Let TheChoice.Letter = PossibleLetters(i).Letter
            Let TheChoice.FileData = PossibleLetters(i).FileData
        End If
    Next i
    If TheChoice.Letter <> "" Then
        Let LetterForWeight = TheChoice.Letter & "  -  Avg Pixel Diff: " & Round((TheChoice.DifferenceValue / 25)) & " pixel(s)"
        For X = 1 To 5
            For Y = 1 To 5
                FillColor X, Y, TheChoice.PixelValue(SectorForGrid(X, Y)), frmOCR.pbDebug2
            Next Y
        Next X
    Else
        If frmOCR.cbEnableAutoTrain Then
            If DrawingOK Then
                Let Status = "Training"
                Let strTemp = UCase(InputBox("What letter is this?"))
                SaveLetter strTemp
                Let LetterForWeight = strTemp
            End If
        End If
    End If
    Let Status = ""
End Function
Private Sub SaveLetter(Letter As String)
    Dim strTmp As String
    Dim X As Long
    Dim Y As Long
    Let strTmp = Letter
    For X = 1 To 5
        For Y = 1 To 5
            Let strTmp = strTmp & "," & Grid(X, Y)
        Next Y
    Next X
    Let strTmp = strTmp & "," & EndY - StartY & "," & EndX - StartX
    Open App.Path & "\LetterInfo.txt" For Append As #1
    Print #1, strTmp
    Close #1
End Sub
Private Sub GatherSizeInformation(ThePic As PictureBox)
    Let Status = "Analyzing image"
    Dim X As Long
    Dim Y As Long
    Let StartY = 0
    Let EndY = 0
    Let StartX = 0
    Let EndX = 0
    Let Progress.Max = 5
    With ThePic
        For Y = 0 To .ScaleHeight - 1
            For X = 0 To .ScaleWidth - 1
                If .Point(X, Y) = 0 Then
                    Let StartY = Y
                    GoTo FoundStartY
                End If
            Next X
        Next Y
    End With
    If StartY = 0 Then Exit Sub
FoundStartY:
    Let Progress.Value = 1
    With ThePic
        For Y = .ScaleHeight - 1 To 0 Step -1
            For X = .ScaleWidth - 1 To 0 Step -1
                If .Point(X, Y) = 0 Then
                    Let EndY = Y
                    GoTo FoundEndY
                End If
            Next X
        Next Y
    End With
    If EndY = 0 Then Exit Sub
FoundEndY:
    Let Progress.Value = 2
    With ThePic
        For X = .ScaleWidth - 1 To 0 Step -1
            For Y = .ScaleHeight - 1 To 0 Step -1
                If .Point(X, Y) = 0 Then
                    Let EndX = X
                    GoTo FoundEndX
                End If
            Next Y
        Next X
    End With
    If EndX = 0 Then Exit Sub
FoundEndX:
    Let Progress.Value = 3
    With ThePic
        For X = 0 To .ScaleWidth - 1
            For Y = 0 To .ScaleHeight - 1
                If .Point(X, Y) = 0 Then
                    Let StartX = X
                    GoTo FoundStartX
                End If
            Next Y
        Next X
    End With
    If StartX = 0 Then Exit Sub
FoundStartX:
    Let Progress.Value = 4
    Let SectionInterval = (EndY - StartY) / 5
    Let XSectionInterval = (EndX - StartX) / 5
    Let Progress.Value = 5
End Sub
