Public Class Form1

    Dim InFileName As String = ""
    Dim ProjectName As String
    Dim LatDispCenter As Double
    Dim LonDispCenter As Double
    Dim Zoom As Integer
    Dim BGLProjectFolder As String

    ' for Maps
    Private Structure Map
        Dim Name As String
        Dim Selected As Boolean
        Dim COLS As Integer ' for the selected map can be one of 0 1 or 2 / not exported!
        Dim ROWS As Integer
        Dim NLAT As Double
        Dim SLAT As Double
        Dim WLON As Double
        Dim ELON As Double
        Dim BMPSu As String
        Dim BMPWi As String
        Dim BMPFa As String
        Dim BMPSp As String
        Dim BMPHw As String
        Dim BMPLm As String
    End Structure
    Dim NoOfMaps As Integer = 0
    Dim Maps() As Map

    ' For Lines
    Private Structure GLine
        Dim Name As String
        Dim Type As String
        Dim Guid As String
        Dim Color As Color
        Dim Selected As Boolean
        Dim NoOfPoints As Integer
        Dim GLPoints() As GLPoint
        'Dim NLAT As Double ' not saved
        'Dim SLAT As Double ' not saved
        'Dim WLON As Double ' not saved
        'Dim ELON As Double ' not saved
        'Dim OnScreen As Boolean
    End Structure
    Dim Lines() As GLine
    Dim NoOfLines As Integer = 0
    Private Structure GLPoint
        Dim lon As Double
        Dim lat As Double
        Dim alt As Double
        Dim wid As Double
        Dim Selected As Boolean
    End Structure

    ' For Polys
    Private Structure GPoly
        Dim Name As String
        Dim Type As String
        Dim Guid As String
        Dim Color As Color
        Dim Selected As Boolean
        Dim NoOfChilds As Integer
        Dim Childs() As Integer
        Dim NoOfPoints As Integer
        Dim GPoints() As GPoint
        'Dim NLAT As Double ' not saved
        'Dim SLAT As Double ' not saved
        'Dim WLON As Double ' not saved
        'Dim ELON As Double ' not saved
        'Dim OnScreen As Boolean
    End Structure
    Private Structure GPoint
        Dim lon As Double
        Dim lat As Double
        Dim alt As Double
        Dim Selected As Boolean
    End Structure
    Dim Polys() As GPoly
    Dim NoOfPolys As Integer

    Private Structure Objecto
        Dim Type As Integer ' 0=FSX 1=FS8 2=FS9 3=Rwy12 4=API 5=ASD 8=TaxiwaySign   128=FSX MDL 129=FS9 MDL
        Dim Description As String ' code the type
        Dim Selected As Boolean
        Dim Width As Single
        Dim Length As Single
        Dim Heading As Single
        Dim Pitch As Single
        Dim Bank As Single
        Dim BiasX As Single
        Dim BiasY As Single
        Dim BiasZ As Single
        Dim lat As Double
        Dim lon As Double
        Dim Altitude As Double
        Dim AGL As Integer
        Dim Complexity As Integer
        ' the following are not exported
        'Dim NLAT As Double
        'Dim SLAT As Double
        'Dim WLON As Double
        'Dim ELON As Double
        'Dim HDX As Single
        'Dim HDY As Single
        'Dim P1Y As Single
        'Dim P1X As Single
        'Dim P2Y As Single
        'Dim P2X As Single
        'Dim P3Y As Single
        'Dim P3X As Single
        'Dim P4Y As Single
        'Dim P4X As Single
    End Structure
    Dim Objects() As Objecto
    Dim NoOfObjects As Integer

    ' For Excludes
    Private Structure Exclude
        Dim Flag As Integer
        Dim Selected As Boolean
        Dim NLAT As Double
        Dim SLAT As Double
        Dim WLON As Double
        Dim ELON As Double
        'Dim CornerSel As Integer ' not saved 1=NW 2=SE 0=none
    End Structure
    Dim Excludes() As Exclude
    Dim NoOfExcludes As Integer


    Private Sub cmdChoose_Click(sender As Object, e As EventArgs) Handles cmdChoose.Click

        OpenFileDialog1.Filter = "SBuilderX Imp/Exp (*.SBX)|*.SBX|All Files|*.*"
        OpenFileDialog1.InitialDirectory = My.Application.Info.DirectoryPath
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Title = "SBuilderX 3.15 - Import Project"

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then 'user did not press cancel
            InFileName = OpenFileDialog1.FileName
            txtFileName.Text = InFileName
        End If
        OpenFileDialog1.Dispose()


    End Sub

    Private Sub cmdConvert_Click(sender As Object, e As EventArgs) Handles cmdConvert.Click

        Dim OutFilename As String

        If InFileName = "" Then Exit Sub

        ImportSBX(InFileName)

        OutFilename = System.IO.Path.GetFileNameWithoutExtension(InFileName) & "_205.SBX"
        OutFilename = System.IO.Path.GetDirectoryName(InFileName) & "\" & OutFilename

        ExportSBX(OutFilename)

        Dim A As String = "Done! Do you want to convert another file?"
        If MsgBox(A, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Dispose()
            Exit Sub

        End If

    End Sub

    Private Sub ImportSBX(ByVal filename As String)

        Dim M, N, J As Integer
        Dim KEY As String

        KEY = ReadIniValue(filename, "Main", "CopyRight")
        KEY = Strings.Left(KEY, 11)

        If KEY <> "PTSIM SB314" Then
            MsgBox("Not a Valid SBX SBuilderX 3.15 File!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        ProjectName = ReadIniValue(filename, "Main", "Name")
        NoOfMaps = ReadIniInteger(filename, "Main", "NoOfMaps")
        NoOfLines = ReadIniInteger(filename, "Main", "NoOfLines")
        NoOfPolys = ReadIniInteger(filename, "Main", "NoOfPolys")

        'NoOfLLXYs = ReadIniInteger(filename, "Main", "NoOfLC_LOD5s")
        'NoOfWWXYs = ReadIniInteger(filename, "Main", "NoOfWC_LOD5s")

        NoOfObjects = ReadIniInteger(filename, "Main", "NoOfObjects")
        NoOfExcludes = ReadIniInteger(filename, "Main", "NoOfExcludes")
        'NoOfLWCIs = ReadIniInteger(filename, "Main", "NoOfLWCIs")

        BGLProjectFolder = ReadIniValue(filename, "Main", "BGLProjectFolder")

        LatDispCenter = ReadIniDouble(filename, "Main", "LatDispCenter")
        LonDispCenter = ReadIniDouble(filename, "Main", "LonDispCenter")

        Zoom = ReadIniInteger(filename, "Main", "Zoom")

        If NoOfMaps > 0 Then ReDim Maps(NoOfMaps)
        If NoOfLines > 0 Then ReDim Lines(NoOfLines)
        If NoOfPolys > 0 Then ReDim Polys(NoOfPolys)
        If NoOfObjects > 0 Then ReDim Objects(NoOfObjects)
        If NoOfExcludes > 0 Then ReDim Excludes(NoOfExcludes)
        'If NoOfLWCIs > 0 Then ReDim LWCIs(NoOfLWCIs)

        For N = 1 To NoOfMaps
            KEY = "Map." & Trim(Str(N))

            Maps(N).Name = ReadIniValue(filename, KEY, "Name")
            Maps(N).BMPSu = ReadIniValue(filename, KEY, "BMPSu")
            Maps(N).BMPSp = ReadIniValue(filename, KEY, "BMPSp")
            Maps(N).BMPFa = ReadIniValue(filename, KEY, "BMPFa")
            Maps(N).BMPWi = ReadIniValue(filename, KEY, "BMPWi")
            Maps(N).BMPHw = ReadIniValue(filename, KEY, "BMPHw")
            Maps(N).BMPLm = ReadIniValue(filename, KEY, "BMPLm")

            Maps(N).COLS = ReadIniInteger(filename, KEY, "Cols")
            Maps(N).ROWS = ReadIniInteger(filename, KEY, "Rows")
            Maps(N).NLAT = ReadIniDouble(filename, KEY, "NLat")
            Maps(N).SLAT = ReadIniDouble(filename, KEY, "SLat")
            Maps(N).ELON = ReadIniDouble(filename, KEY, "ELon")
            Maps(N).WLON = ReadIniDouble(filename, KEY, "WLon")

        Next N

        FileOpen(5, filename, OpenMode.Input)

        For N = 1 To NoOfLines
            KEY = "[Line." & Trim(Str(N)) & "]"
            GoToThisKey((KEY))
            Lines(N).Name = Mid(LineInput(5), 6)
            Lines(N).Type = Mid(LineInput(5), 6)
            Lines(N).Guid = Mid(LineInput(5), 6)
            Lines(N).Color = ColorFromArgb(Mid(LineInput(5), 7))
            Lines(N).NoOfPoints = CInt(Mid(LineInput(5), 12))
            If Lines(N).Name = "" Then Lines(N).Name = CStr(Lines(N).NoOfPoints) & "_Pts_Imported_Line"
            ReDim Lines(N).GLPoints(Lines(N).NoOfPoints)
            J = 6
            For M = 1 To Lines(N).NoOfPoints
                If M > 9 Then J = 7
                If M > 99 Then J = 8
                If M > 999 Then J = 9
                If M > 9999 Then J = 10
                If M > 99999 Then J = 11
                If M > 999999 Then J = 12
                Lines(N).GLPoints(M).lat = Val(Mid(LineInput(5), J))
                Lines(N).GLPoints(M).lon = Val(Mid(LineInput(5), J))
                Lines(N).GLPoints(M).alt = Val(Mid(LineInput(5), J))
                Lines(N).GLPoints(M).wid = Val(Mid(LineInput(5), J))
            Next M

        Next N

        For N = 1 To NoOfPolys
            KEY = "[Poly." & Trim(Str(N)) & "]"
            GoToThisKey((KEY))
            Polys(N).Name = Mid(LineInput(5), 6)
            Polys(N).Type = Mid(LineInput(5), 6)
            Polys(N).Guid = Mid(LineInput(5), 6)
            Polys(N).Color = ColorFromArgb(Mid(LineInput(5), 7))
            J = CInt(Mid(LineInput(5), 12))
            Polys(N).NoOfChilds = J
            If J > 0 Then
                ReDim Polys(N).Childs(J)
            Else
                ReDim Polys(N).Childs(0)
            End If
            J = 8
            For M = 1 To Polys(N).NoOfChilds
                If M > 9 Then J = 9
                If M > 99 Then J = 10
                If M > 999 Then J = 11
                If M > 9999 Then J = 12
                If M > 99999 Then J = 13
                If M > 999999 Then J = 14
                Polys(N).Childs(M) = Val(Mid(LineInput(5), J))
            Next
            Polys(N).NoOfPoints = CInt(Mid(LineInput(5), 12))
            If Polys(N).Name = "" Then Polys(N).Name = CStr(Polys(N).NoOfPoints) & "_Pts_Imported_Polygon"
            ReDim Polys(N).GPoints(Polys(N).NoOfPoints)
            J = 6
            For M = 1 To Polys(N).NoOfPoints
                If M > 9 Then J = 7
                If M > 99 Then J = 8
                If M > 999 Then J = 9
                If M > 9999 Then J = 10
                If M > 99999 Then J = 11
                If M > 999999 Then J = 12
                Polys(N).GPoints(M).lat = Val(Mid(LineInput(5), J))
                Polys(N).GPoints(M).lon = Val(Mid(LineInput(5), J))
                Polys(N).GPoints(M).alt = Val(Mid(LineInput(5), J))

            Next M
        Next N

        For N = 1 To NoOfExcludes
            KEY = "[Exclude." & Trim(Str(N)) & "]"
            GoToThisKey((KEY))
            Excludes(N).Flag = CInt(Mid(LineInput(5), 6))
            Excludes(N).NLAT = Val(Mid(LineInput(5), 6))
            Excludes(N).SLAT = Val(Mid(LineInput(5), 6))
            Excludes(N).ELON = Val(Mid(LineInput(5), 6))
            Excludes(N).WLON = Val(Mid(LineInput(5), 6))
        Next N


        For N = 1 To NoOfObjects
            KEY = "[Object." & Trim(Str(N)) & "]"
            GoToThisKey((KEY))
            Objects(N).Type = CInt(Mid(LineInput(5), 6))
            Objects(N).Description = Mid(LineInput(5), 13)
            Objects(N).Width = CSng(Val(Mid(LineInput(5), 7)))
            Objects(N).Length = CSng(Val(Mid(LineInput(5), 8)))
            Objects(N).Heading = CSng(Val(Mid(LineInput(5), 9)))
            Objects(N).Pitch = CSng(Val(Mid(LineInput(5), 7)))
            Objects(N).Bank = CSng(Val(Mid(LineInput(5), 6)))
            Objects(N).BiasX = CSng(Val(Mid(LineInput(5), 7)))
            Objects(N).BiasY = CSng(Val(Mid(LineInput(5), 7)))
            Objects(N).BiasZ = CSng(Val(Mid(LineInput(5), 7)))
            Objects(N).lat = Val(Mid(LineInput(5), 5))
            Objects(N).lon = Val(Mid(LineInput(5), 5))
            Objects(N).Altitude = CSng(Val(Mid(LineInput(5), 10)))
            Objects(N).AGL = CInt(Mid(LineInput(5), 5))
            Objects(N).Complexity = CInt(Mid(LineInput(5), 12))
            ' AddLatLonToObjects(N)

        Next N

        FileClose(5)


    End Sub

    Private Function ReadIniValue(ByVal INIpath As String, ByVal KEY As String, ByRef Variable As String) As String

        Dim NF As Integer
        Dim Temp As String
        Dim LcaseTemp As String
        Dim ReadyToRead As Boolean

AssignVariables:
        NF = FreeFile()
        ReadIniValue = ""
        KEY = "[" & LCase(KEY) & "]"
        Variable = LCase(Variable)

EnsureFileExists:
        FileOpen(NF, INIpath, OpenMode.Binary)
        FileClose(NF)
        SetAttr(INIpath, FileAttribute.Archive)

LoadFile:
        FileOpen(NF, INIpath, OpenMode.Input)
        While Not EOF(NF)
            Temp = LineInput(NF)
            LcaseTemp = LCase(Temp)
            If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
            If LcaseTemp = KEY Then ReadyToRead = True
            If InStr(LcaseTemp, "[") = 0 And ReadyToRead Then
                If InStr(LcaseTemp, Variable & "=") = 1 Then
                    ReadIniValue = Mid(Temp, 1 + Len(Variable & "="))
                    FileClose(NF)
                    Exit Function
                End If
            End If
        End While
        FileClose(NF)
    End Function

    Private Function ReadIniInteger(ByVal File As String, ByVal KEY As String, ByVal Value As String) As Integer

        On Error GoTo erro
        ReadIniInteger = CInt(ReadIniValue(File, KEY, Value))
        Exit Function
erro:
        ReadIniInteger = 0

    End Function

    Private Function ReadIniDouble(ByVal File As String, ByVal KEY As String, ByVal Value As String) As Double

        On Error GoTo erro
        ReadIniDouble = Val(ReadIniValue(File, KEY, Value))
        Exit Function
erro:
        ReadIniDouble = 0

    End Function

    Private Sub GoToThisKey(ByVal KEY As String)

        Dim A As String
        Dim NK, NA As Integer

        NK = Len(KEY)

        Do
            A = LineInput(5)
            NA = Len(A)
            If NA = NK Then If A = KEY Then Exit Do
        Loop

    End Sub

    Private Function ColorFromArgb(ByVal argb As String) As Color

        ColorFromArgb = Color.FromArgb(Convert.ToInt32(argb, 16))

    End Function

    Private Function IntegerFromColor(ByVal myColor As Color) As Integer

        IntegerFromColor = RGB(myColor.R, myColor.G, myColor.B)

    End Function


    Private Sub ExportSBX(ByVal FileName As String)

        Dim M, N, FN As Integer
        Dim KEY As String

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        FN = FreeFile()

        FileOpen(FN, FileName, OpenMode.Output)

        PrintLine(FN, "[Main]")
        PrintLine(FN, "CopyRight=" & "PTSIM SB205")
        PrintLine(FN, "Name=" & ProjectName)

        PrintLine(FN, "NoOfPhotos=0")
        PrintLine(FN, "NoOfMaps=" & NoOfMaps)
        PrintLine(FN, "NoOfLands=0")
        PrintLine(FN, "NoOfLines=" & NoOfLines)
        PrintLine(FN, "NoOfPolys=" & NoOfPolys)
        PrintLine(FN, "NoOfWaters=0")
        PrintLine(FN, "NoOfObjects=" & NoOfObjects)
        PrintLine(FN, "NoOfFills=0")
        PrintLine(FN, "NoOfExcludes=" & NoOfExcludes)
        PrintLine(FN, "NoOfMeshes=0")
        PrintLine(FN, "NoOfLWCIs=0")

        PrintLine(FN, "PhotoSeasons=1281")

        PrintLine(FN, "BGL1ProjectFolder=" & BGLProjectFolder)
        PrintLine(FN, "BGL2ProjectFolder=" & BGLProjectFolder)

        PrintLine(FN, "LatDispCenter=" & Str(LatDispCenter))
        PrintLine(FN, "LonDispCenter=" & Str(LonDispCenter))

        PrintLine(FN, "LatLandCenter=" & Str(LatDispCenter))
        PrintLine(FN, "LonLandCenter=" & Str(LonDispCenter))
        PrintLine(FN, "ClassGridVIEW=0")

        PrintLine(FN, "Zoom=" & Str(2 ^ (Zoom - 9)))

        For N = 1 To NoOfMaps

            PrintLine(FN)
            KEY = "[Map." & Trim(Str(N)) & "]"
            PrintLine(FN, KEY)
            PrintLine(FN, "Name=" & Maps(N).Name)
            PrintLine(FN, "BMPSu=" & Maps(N).BMPSu)
            PrintLine(FN, "BMPSp=" & Maps(N).BMPSp)
            PrintLine(FN, "BMPFa=" & Maps(N).BMPFa)
            PrintLine(FN, "BMPWi=" & Maps(N).BMPWi)
            PrintLine(FN, "BMPHw=" & Maps(N).BMPHw)
            PrintLine(FN, "BMPLm=" & Maps(N).BMPLm)
            ' repeat as I do not remeber what it was used for
            PrintLine(FN, "BMPAl=" & Maps(N).BMPSu)
            PrintLine(FN, "BMPLC=" & Maps(N).BMPSu)
            PrintLine(FN, "BMPME=" & Maps(N).BMPSu)

            PrintLine(FN, "COLS0=" & Str(Maps(N).COLS))
            PrintLine(FN, "ROWS0=" & Str(Maps(N).ROWS))
            PrintLine(FN, "NLat0=" & Str(Maps(N).NLAT))
            PrintLine(FN, "SLat0=" & Str(Maps(N).SLAT))
            PrintLine(FN, "ELon0=" & Str(Maps(N).ELON))
            PrintLine(FN, "WLon0=" & Str(Maps(N).WLON))
            ' repeat as I do not remeber what it was used for
            PrintLine(FN, "COLS1=" & Str(Maps(N).COLS))
            PrintLine(FN, "ROWS1=" & Str(Maps(N).ROWS))
            PrintLine(FN, "NLat1=" & Str(Maps(N).NLAT))
            PrintLine(FN, "SLat1=" & Str(Maps(N).SLAT))
            PrintLine(FN, "ELon1=" & Str(Maps(N).ELON))
            PrintLine(FN, "WLon1=" & Str(Maps(N).WLON))
            ' repeat as I do not remeber what it was used for
            PrintLine(FN, "COLS2=" & Str(Maps(N).COLS))
            PrintLine(FN, "ROWS2=" & Str(Maps(N).ROWS))
            PrintLine(FN, "NLat2=" & Str(Maps(N).NLAT))
            PrintLine(FN, "SLat2=" & Str(Maps(N).SLAT))
            PrintLine(FN, "ELon2=" & Str(Maps(N).ELON))
            PrintLine(FN, "WLon2=" & Str(Maps(N).WLON))

        Next N

        For N = 1 To NoOfLines
            PrintLine(FN)
            KEY = "[Line." & Trim(Str(N)) & "]"
            PrintLine(FN, KEY)
            PrintLine(FN, "Name=" & Lines(N).Name)
            PrintLine(FN, "Type= ")
            'PrintLine(FN, "Guid=" & Lines(N).Guid)
            PrintLine(FN, "Color=" & IntegerFromColor(Lines(N).Color))
            PrintLine(FN, "NoOfPoints=" & Str(Lines(N).NoOfPoints))
            For M = 1 To Lines(N).NoOfPoints
                PrintLine(FN, "Lat" & Trim(Str(M)) & "=" & Str(Lines(N).GLPoints(M).lat))
                PrintLine(FN, "Lon" & Trim(Str(M)) & "=" & Str(Lines(N).GLPoints(M).lon))
                'PrintLine(FN, "Alt" & M & "=" & Str(Lines(N).GLPoints(M).alt))
                PrintLine(FN, "Wid" & M & "=" & Str(Lines(N).GLPoints(M).wid))
            Next M
        Next N

        For N = 1 To NoOfPolys
            PrintLine(FN)
            KEY = "[Poly." & Trim(Str(N)) & "]"
            PrintLine(FN, KEY)
            PrintLine(FN, "Name=" & Polys(N).Name)
            PrintLine(FN, "Type=")
            ' PrintLine(FN, "Guid=" & Polys(N).Guid)
            PrintLine(FN, "Color=" & IntegerFromColor(Polys(N).Color))
            'PrintLine(FN, "NoOfChilds=" & CStr(Polys(N).NoOfChilds))
            'For M = 1 To Polys(N).NoOfChilds
            '    PrintLine(FN, "Child" & Trim(Str(M)) & "=" & Str(Polys(N).Childs(M)))
            'Next M
            PrintLine(FN, "NoOfPoints=" & Str(Polys(N).NoOfPoints))
            For M = 1 To Polys(N).NoOfPoints
                PrintLine(FN, "Lat" & Trim(Str(M)) & "=" & Str(Polys(N).GPoints(M).lat))
                PrintLine(FN, "Lon" & Trim(Str(M)) & "=" & Str(Polys(N).GPoints(M).lon))
                PrintLine(FN, "Alt" & Trim(Str(M)) & "=" & Str(Polys(N).GPoints(M).alt))
            Next M
        Next N

        For N = 1 To NoOfExcludes
            PrintLine(FN)
            KEY = "[Exclude." & Trim(Str(N)) & "]"
            PrintLine(FN, KEY)
            PrintLine(FN, "Flag=" & Str(Excludes(N).Flag))
            PrintLine(FN, "NLat=" & Str(Excludes(N).NLAT))
            PrintLine(FN, "SLat=" & Str(Excludes(N).SLAT))
            PrintLine(FN, "ELon=" & Str(Excludes(N).ELON))
            PrintLine(FN, "WLon=" & Str(Excludes(N).WLON))
        Next N

        For N = 1 To NoOfObjects
            PrintLine(FN)
            KEY = "[Object." & Trim(Str(N)) & "]"
            PrintLine(FN, KEY)

            PrintLine(FN, "Type=" & Str(Objects(N).Type))
            PrintLine(FN, "Description=" & Objects(N).Description)
            PrintLine(FN, "Width=" & Str(Objects(N).Width))
            PrintLine(FN, "Length=" & Str(Objects(N).Length))
            PrintLine(FN, "Heading=" & Str(Objects(N).Heading))
            PrintLine(FN, "Pitch=" & Str(Objects(N).Pitch))
            PrintLine(FN, "Bank=" & Str(Objects(N).Bank))
            PrintLine(FN, "BiasX=" & Str(Objects(N).BiasX))
            PrintLine(FN, "BiasY=" & Str(Objects(N).BiasY))
            PrintLine(FN, "BiasZ=" & Str(Objects(N).BiasZ))
            PrintLine(FN, "Lat=" & Str(Objects(N).lat))
            PrintLine(FN, "Lon=" & Str(Objects(N).lon))
            PrintLine(FN, "Altitude=" & Str(Objects(N).Altitude))
            PrintLine(FN, "AGL=" & Str(Objects(N).AGL))
            PrintLine(FN, "Complexity=" & Str(Objects(N).Complexity))
        Next N

        FileClose(FN)

        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        Dispose()

    End Sub


End Class
