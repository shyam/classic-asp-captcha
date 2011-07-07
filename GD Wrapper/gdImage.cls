VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gdImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'' This file was writen by Trevor Herselman

Dim clsImagePtr As Long
Dim clsStream As ADODB.Stream


' Wrapper functions
' =================


Public Type gdPoint
    x As Long
    y As Long
End Type


Private Sub Class_Terminate()
    If clsImagePtr <> 0 Then gdImageDestroy (clsImagePtr)
    If TypeName(clsStream) = "Stream" Then If clsStream.State <> adStateClosed Then clsStream.Close
    Set clsStream = Nothing
End Sub

Public Property Get Image() As Long
    Image = clsImagePtr
End Property
Public Property Let Image(newImage As Long)
    If clsImagePtr <> 0 Then gdImageDestroy (clsImagePtr)
    clsImagePtr = newImage '' Used for some functions that return a completely NEW image!
End Property
Public Property Get Stream() As ADODB.Stream
    Set Stream = clsStream
End Property

Public Sub Create(width As Long, height As Long)
    clsImagePtr = gdImageCreate(width, height)
End Sub
Public Sub CreateTrueColor(width As Long, height As Long)
    clsImagePtr = gdImageCreateTrueColor(width, height)
End Sub
Public Sub Destroy()
    If clsImagePtr <> 0 Then gdImageDestroy (clsImagePtr)
    clsImagePtr = 0
    If TypeName(clsStream) = "Stream" Then clsStream.Close
    Set clsStream = Nothing
End Sub

Public Function ColorAllocate(Red As Long, Green As Long, Blue As Long) As Long
    ColorAllocate = gdImageColorAllocate(clsImagePtr, Red, Green, Blue)
End Function


Public Sub TrueColorToPalette(Optional ditherFlag As Long = False, Optional colorsWanted As Long = 256)
    Call gdImageTrueColorToPalette(clsImagePtr, ditherFlag, colorsWanted)
End Sub
Public Function CreatePaletteFromTrueColor(Optional ditherFlag As Long = False, Optional colorsWanted As Long = 256) As Long
    CreatePaletteFromTrueColor = gdImageCreatePaletteFromTrueColor(clsImagePtr, ditherFlag, colorsWanted)
End Function


Public Sub SetPixel(x As Long, y As Long, color As Long)
    Call gdImageSetPixel(clsImagePtr, x, y, color)
End Sub
Public Function GetPixel(x As Long, y As Long) As Long
    GetPixel = gdImageGetPixel(clsImagePtr, x, y)
End Function


' Drawing Functions
' =================


Public Sub Line(x1 As Long, y1 As Long, x2 As Long, y2 As Long, color As Long)
    Call gdImageLine(clsImagePtr, x1, y1, x2, y2, color)
End Sub
Public Sub Polygon(gdPoints() As gdPoint, pointsTotal As Long, color As Long)
    Call gdImagePolygon(clsImagePtr, gdPoints(0), pointsTotal, color)
End Sub
Public Sub OpenPolygon(gdPoints() As gdPoint, pointsTotal As Long, color As Long)
    Call gdImageOpenPolygon(clsImagePtr, gdPoints(0), pointsTotal, color)
End Sub
Public Sub Rectangle(x1 As Long, y1 As Long, x2 As Long, y2 As Long, color As Long)
    Call gdImageRectangle(clsImagePtr, x1, y1, x2, y2, color)
End Sub
Public Sub FilledPolygon(gdPoints() As gdPoint, pointsTotal As Long, color As Long)
    Call gdImageFilledPolygon(clsImagePtr, gdPoints(0), pointsTotal, color)
End Sub
Public Sub FilledRectangle(x1 As Long, y1 As Long, x2 As Long, y2 As Long, color As Long)
    Call gdImageFilledRectangle(clsImagePtr, x1, y1, x2, y2, color)
End Sub
Public Sub Arc(x As Long, y As Long, width As Long, height As Long, degreeStart As Long, degreeEnd As Long, color As Long)
    Call gdImageArc(clsImagePtr, x, y, width, height, degreeStart, degreeEnd, color)
End Sub
Public Sub Ellipse(x As Long, y As Long, width As Long, height As Long, color As Long) '' My own wrapper function!
    Call gdImageArc(clsImagePtr, x, y, width, height, 0, 360, color)
End Sub
Public Sub FilledArc(x As Long, y As Long, width As Long, height As Long, degreeStart As Long, degreeEnd As Long, color As Long, Optional style As Long = gdArc)
    Call gdImageFilledArc(clsImagePtr, x, y, width, height, degreeStart, degreeEnd, color, style)
End Sub
Public Sub FilledEllipse(x As Long, y As Long, width As Long, height As Long, color As Long)
    Call gdImageFilledEllipse(clsImagePtr, x, y, width, height, color)
End Sub
Public Sub FillToBorder(x As Long, y As Long, borderColor As Long, color As Long)
    Call gdImageFillToBorder(clsImagePtr, x, y, borderColor, color)
End Sub
Public Sub Fill(x As Long, y As Long, color As Long)
    Call gdImageFill(clsImagePtr, x, y, color)
End Sub
Public Sub SetAntiAliased(c As Long)
    Call gdImageSetAntiAliased(clsImagePtr, c)
End Sub
Public Sub SetAntiAliasedDontBlend(c As Long)
    Call gdImageSetAntiAliasedDontBlend(clsImagePtr, c)
End Sub
Public Sub SetBrush(brush As Long)
    Call gdImageSetBrush(clsImagePtr, brush)
End Sub
Public Sub SetTile(tile As Long)
    Call gdImageSetTile(clsImagePtr, tile)
End Sub
Public Sub SetStyle(style As Long, styleLength As Long)
    Call gdImageSetStyle(clsImagePtr, style, styleLength)
End Sub
Public Sub SetThickness(thickness As Long)
    Call gdImageSetThickness(clsImagePtr, thickness)
End Sub
Public Sub AlphaBlending(blending As Long)
    Call gdImageAlphaBlending(clsImagePtr, blending)
End Sub
Public Sub SaveAlpha(saveFlag As Long)
    Call gdImageSaveAlpha(clsImagePtr, saveFlag)
End Sub
Public Sub SetClip(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    Call gdImageSetClip(clsImagePtr, x1, y1, x2, y2)
End Sub
Public Sub GetClip(x1P As Long, y1P As Long, x2P As Long, y2P As Long)
    Call gdImageGetClip(clsImagePtr, x1P, y1P, x2P, y2P)
End Sub


' Font Functions
' ==============


Public Function FontGetSmall() As Long
    FontGetSmall = gdFontGetSmall()
End Function
Public Function FontGetLarge() As Long
    FontGetLarge = gdFontGetLarge()
End Function
Public Function FontGetMediumBold() As Long
    FontGetMediumBold = gdFontGetMediumBold()
End Function
Public Function FontGetGiant() As Long
    FontGetGiant = gdFontGetGiant()
End Function
Public Function FontGetTiny() As Long
    FontGetTiny = gdFontGetTiny()
End Function

Public Sub Char(font As Long, x As Long, y As Long, Character As Long, color As Long)
    Call gdImageChar(clsImagePtr, font, x, y, Character, color)
End Sub
Public Sub CharUp(font As Long, x As Long, y As Long, Character As Long, color As Long)
    Call gdImageCharUp(clsImagePtr, font, x, y, Character, color)
End Sub
Public Sub Chars(font As Long, x As Long, y As Long, Characters As String, color As Long)
    Call gdImageString(clsImagePtr, font, x, y, Characters, color)
End Sub
Public Sub CharsUp(font As Long, x As Long, y As Long, Characters As String, color As Long)
    Call gdImageStringUp(clsImagePtr, font, x, y, Characters, color)
End Sub

'Public Function FTUseFontConfig(flag As Long) As Long
'    FTUseFontConfig = gdFTUseFontConfig(flag)
'End Function
'Public Function StringFT(brect() As Long, color As Long, fontname As String, ptsize As Double, angle As Double, x As Long, y As Long, text As String) As String
'    StringFT = gdImageStringFT(clsImagePtr, brect(), color, fontname, ptsize, angle, x, y, text)
'End Function


' MACROS
' ======


Public Function width() As Long
    Call CopyMemory(width, ByVal clsImagePtr + 4, 4)
End Function
Public Function height() As Long
    Call CopyMemory(height, ByVal clsImagePtr + 8, 4)
End Function
Public Function ColorsTotal() As Long
    '' AKA gdImageColorsTotal MACRO
    Call CopyMemory(ColorsTotal, ByVal clsImagePtr + 12, 4)
End Function

Public Function Red(Index As Long) As Long
    '' AKA gdImageRed MACRO
    Call CopyMemory(Red, ByVal clsImagePtr + 16 + Index, 4)
End Function
'Public Function Green(Index As Long) As Long
'    '' AKA gdImageGreen MACRO
'    Call CopyMemory(Green, ByVal clsImagePtr + ??? + Index, 4)
'End Function
'Public Function Blue(Index As Long) As Long
'    '' AKA gdImageBlue MACRO
'    Call CopyMemory(Blue, ByVal clsImagePtr + ??? + Index, 4)
'End Function





'' Client will be responsible for releasing the memory with Free()
Public Function ToGifPtr(size As Long) As Long
    ToGifPtr = gdImageGifPtr(clsImagePtr, size)
End Function
Public Function ToPngPtr(size As Long) As Long
    ToPngPtr = gdImagePngPtr(clsImagePtr, size)
End Function
Public Function ToJpegPtr(size As Long) As Long
    ToJpegPtr = gdImageJpegPtr(clsImagePtr, size)
End Function
Public Sub Free(gdPtr As Long)
    Call gdFree(gdPtr)
End Sub


' Extended functions
' ==================


Public Function color(Red As Integer, Green As Integer, Blue As Integer) As Long
    color = RGB(Red, Green, Blue)
End Function
Public Function GetRed(RGBA As Long) As Long
    Dim Colors(2) As Byte
    Call CopyMemory(Colors(0), RGBA, 3)
    GetRed = Colors(0)
End Function
Public Function GetGreen(RGBA As Long) As Long
    Dim Colors(2) As Byte
    Call CopyMemory(Colors(0), RGBA, 3)
    GetGreen = Colors(1)
End Function
Public Function GetBlue(RGBA As Long) As Long
    Dim Colors(2) As Byte
    Call CopyMemory(Colors(0), RGBA, 3)
    GetBlue = Colors(2)
End Function


Public Sub GradientFillRect(FromColor As Long, ToColor As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional Vertical As Boolean = False)
    Dim FromRGB(2) As Byte, ToRGB(2) As Byte, Counter As Long, Percent As Double, DiffRGB(2) As Long

    Call CopyMemory(FromRGB(0), FromColor, 3)
    Call CopyMemory(ToRGB(0), ToColor, 3)

    DiffRGB(0) = CLng(ToRGB(0)) - CLng(FromRGB(0))
    DiffRGB(1) = CLng(ToRGB(1)) - CLng(FromRGB(1))
    DiffRGB(2) = CLng(ToRGB(2)) - CLng(FromRGB(2))

    Select Case Vertical
        Case False
            For Counter = x1 To x2
                Let Percent = ((Counter - x1) / (x2 - x1)) * 100
                Call Line(Counter, y1, Counter, y2, ColorAllocate(FromRGB(0) + ((DiffRGB(0) / 100) * Percent), FromRGB(1) + ((DiffRGB(1) / 100) * Percent), FromRGB(2) + ((DiffRGB(2) / 100) * Percent)))
            Next
        Case True
            For Counter = y1 To y2
                Let Percent = ((Counter - y1) / (y2 - y1)) * 100
                Call Line(x1, Counter, x2, Counter, ColorAllocate(FromRGB(0) + ((DiffRGB(0) / 100) * Percent), FromRGB(1) + ((DiffRGB(1) / 100) * Percent), FromRGB(2) + ((DiffRGB(2) / 100) * Percent)))
            Next
    End Select
End Sub


Public Function LoadFromFile(FileName As String) As Boolean
    If clsImagePtr <> 0 Then gdImageDestroy (clsImagePtr)
    clsImagePtr = 0 '' Could crash on loading file!
    If TypeName(clsStream) = "Stream" Then
        If clsStream.State <> adStateClosed Then clsStream.Close
    Else
        Set clsStream = New ADODB.Stream
    End If

    clsStream.Open
    clsStream.Type = adTypeBinary
On Error Resume Next
    clsStream.LoadFromFile (FileName)
On Error GoTo 0
    If clsStream.size > 0 Then
        Dim data() As Byte
        data = clsStream.Read
        clsStream.Position = 0

        Select Case LCase(Right(FileName, Len(FileName) - InStrRev(FileName, ".") + 1))
            '' AKA Select Case Extention
            Case ".gif": clsImagePtr = gdImageCreateFromGifPtr(UBound(data) + 1, data(0))
            Case ".png": clsImagePtr = gdImageCreateFromPngPtr(UBound(data) + 1, data(0))
            Case ".jpg", ".jpeg": clsImagePtr = gdImageCreateFromJpegPtr(UBound(data) + 1, data(0))
            Case ".bmp":  clsImagePtr = gdImageCreateFromWBMPPtr(UBound(data) + 1, data(0))
            Case ".gd": clsImagePtr = gdImageCreateFromGdPtr(UBound(data) + 1, data(0))
            Case ".gd2": clsImagePtr = gdImageCreateFromGd2Ptr(UBound(data) + 1, data(0))
            'Case ".xbm": clsImagePtr = gdImageCreateFromXpm(FileName)
        End Select
        
        LoadFromFile = True
    End If
End Function
Public Sub SaveToFile(FileName As String, Optional Options As SaveOptionsEnum = adSaveCreateNotExist)
    If clsImagePtr = 0 Then Exit Sub

    Dim gdPtr As Long
    Dim size As Long
    Select Case LCase(Right(FileName, Len(FileName) - InStrRev(FileName, ".") + 1))
        '' AKA Select Case Extention
        Case ".gif": gdPtr = gdImageGifPtr(clsImagePtr, size)
        Case ".png": gdPtr = gdImagePngPtrEx(clsImagePtr, size)
        Case ".jpg", ".jpeg": gdPtr = gdImageJpegPtr(clsImagePtr, size)
        Case ".bmp": gdPtr = gdImageWBMPPtr(clsImagePtr, size)
        Case ".gd": gdPtr = gdImageGdPtr(clsImagePtr, size)
    End Select

    '' Copy memory from the allocated buffer, into a Byte() array for the Stream objects (requires a Byte() array!)
    Dim buffer() As Byte
    ReDim buffer(size - 1) As Byte
    Call CopyMemory(buffer(0), ByVal gdPtr, ByVal size)
    gdFree (gdPtr)

    '' Reset the stream object
    If TypeName(clsStream) = "Stream" Then
        If clsStream.State <> adStateClosed Then clsStream.Close
    Else
        Set clsStream = New ADODB.Stream
    End If

    '' Write the Byte() array to the stream object
    Call clsStream.Open
    clsStream.Type = adTypeBinary
    Call clsStream.Write(buffer)
    clsStream.Position = 0

    '' Save the file
On Error Resume Next
    Call clsStream.SaveToFile(FileName, Options)
On Error GoTo 0
End Sub


'Public Function SaveAs(Filename As String) As Long
'    Dim hFile As Long
'    Dim gdPtr As Long
'    Dim size As Long
'    Dim BytesWritten As Long
'
'    hFile = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
'
'    Select Case LCase(Right(Filename, Len(Filename) - InStrRev(Filename, ".") + 1))
'        '' AKA Select Case Extention
'        Case ".gif": gdPtr = gdImageGifPtr(clsImagePtr, size)
'        Case ".png": 'gdPtr = gdImagePngPtr(gdImagePtr, size)
'        Case ".jpg": 'gdPtr = gdImageGifPtr(gdImagePtr, size)
'        Case ".bmp": 'gdPtr = gdImageGifPtr(gdImagePtr, size)
'    End Select
'
'    Call WriteFile(ByVal hFile, ByVal gdPtr, ByVal size, BytesWritten)
'    Call CloseHandle(hFile)
'    gdFree (gdPtr)
'
'    SaveAs = BytesWritten
'End Function
'
'
'Public Function SavePtrAs(ByVal gdPtr As Long, ByVal size As Long, Filename As String) As Long
'    Dim hFile As Long
'    Dim BytesWritten As Long
'
'    hFile = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
'    Call WriteFile(ByVal hFile, ByVal gdPtr, ByVal size, BytesWritten)
'    Call CloseHandle(hFile)
'
'    SavePtrAs = BytesWritten
'End Function
'
'
'Public Function ToStream(gdPtr As Long, size As Long) As ADODB.Stream
'    Dim buffer() As Byte
'    ReDim buffer(size)
'    Call CopyMemory(buffer(0), ByVal gdPtr, size)
'
'    Set ToStream = New ADODB.Stream
'    Call ToStream.Open
'    ToStream.Type = adTypeBinary
'    Call ToStream.Write(buffer)
'
'    ToStream.Position = 0
'End Function

Public Function ToGifStream() As ADODB.Stream
    Dim gdPtr As Long
    Dim size As Long
    gdPtr = gdImageGifPtr(ByVal clsImagePtr, size)
    
    Dim buffer() As Byte
    ReDim buffer(size - 1) As Byte
    Call CopyMemory(buffer(0), ByVal gdPtr, ByVal size)
    gdFree (gdPtr)

    Set ToGifStream = New ADODB.Stream
    Set clsStream = ToGifStream
    Call ToGifStream.Open
    ToGifStream.Type = adTypeBinary
    Call ToGifStream.Write(buffer)

    ToGifStream.Position = 0
End Function
Public Function ToPngStream() As ADODB.Stream
    Dim gdPtr As Long
    Dim size As Long
    gdPtr = gdImagePngPtrEx(ByVal clsImagePtr, size)
    
    Dim buffer() As Byte
    ReDim buffer(size - 1) As Byte
    Call CopyMemory(buffer(0), ByVal gdPtr, ByVal size)
    gdFree (gdPtr)

    Set ToPngStream = New ADODB.Stream
    Set clsStream = ToPngStream
    Call ToPngStream.Open
    ToPngStream.Type = adTypeBinary
    Call ToPngStream.Write(buffer)

    ToPngStream.Position = 0
End Function
Public Function ToJpegStream() As ADODB.Stream
    Dim gdPtr As Long
    Dim size As Long
    gdPtr = gdImageJpegPtr(ByVal clsImagePtr, size, 100)
    
    Dim buffer() As Byte
    ReDim buffer(size - 1) As Byte
    Call CopyMemory(buffer(0), ByVal gdPtr, ByVal size)
    gdFree (gdPtr)

    Set ToJpegStream = New ADODB.Stream
    Set clsStream = ToJpegStream
    Call ToJpegStream.Open
    ToJpegStream.Type = adTypeBinary
    Call ToJpegStream.Write(buffer)

    ToJpegStream.Position = 0
End Function
'Public Function ToBmpStream() As ADODB.Stream
'    Dim gdPtr As Long
'    Dim size As Long
'    gdPtr = gdImageWBMPPtr(ByVal clsImagePtr, size)
'
'    Dim buffer() As Byte
'    ReDim buffer(size - 1) As Byte
'    Call CopyMemory(buffer(0), ByVal gdPtr, ByVal size)
'    gdFree (gdPtr)
'
'    Set ToBmpStream = New ADODB.Stream
'    Set clsStream = ToBmpStream
'    Call ToBmpStream.Open
'    ToBmpStream.Type = adTypeBinary
'    Call ToBmpStream.Write(buffer)
'
'    ToBmpStream.Position = 0
'End Function
