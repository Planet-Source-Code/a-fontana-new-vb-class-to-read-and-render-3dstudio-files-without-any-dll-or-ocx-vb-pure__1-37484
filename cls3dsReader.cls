VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3dsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************* '
'                                                           '
'     Date:   16/07/2002                                    '
'    Class:   cls3dsReader                                  '
'   Author:   Andrea Fontana                                '
'   E-Mail:   trikko@katamail.com                           '
' Homepage:   www.it.owns.it - www.vbp.it/trikko            '
'                                                           '
' ********************************************************* '

Option Explicit  '' All vars must be declared.

''-----------------------------
''   Types Declarations
''------------------------------

Private Type RECT         ''
    Left As Long          ''
    Top As Long           ''   API's Rect-Type
    Right As Long         ''
    Bottom As Long        ''
End Type

Private Type POINTAPI     ''
    X As Single           ''   API's POINTAPI-Type
    Y As Single           ''
End Type

Private Type tPnt         ''
    X As Single           ''
    Y As Single           ''   Three-Dimensions Coord.
    Z As Single           ''
End Type                  ''

Private Type tChunk       ''
    Header As Integer     ''   3ds' files are make of chunk.
    Length As Long        ''   Every chunk has a header, a
End Type                  ''   lenght and contain some infos.
                          ''
                          
Private Type tSeg         ''
    Points() As Integer   ''   Segment-Type
    nPoints As Long       ''
End Type

Private Type tSolid       ''
    Points() As tPnt      ''
    nPoints As Integer    ''   Solid-Type
    nSegs As Integer      ''
    Segs() As tSeg        ''
End Type


''-----------------------------
''   API Declarations >>>
''   (these functions provide
''    to draw the solid)
''------------------------------

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

''-----------------------------
''      Consts
''          and
''            Vars
''------------------------------

                          
                              ''  Objects' Vars:
Private nnPoints As Long      ''  >> Temporary number of points
Private nnSegs As Long        ''  >> Temporary number of segments
Private tnSolids As Long      ''  >> Total number of solids
Private tnPoints As Long      ''  >> Total number of points
Private tnPolys As Long       ''  >> Total number of polygons

                              ''  Other Infos:
Private s3dStudioV As String  ''  >> 3dStudio's Version.
Private sMeshV As String      ''  >> Meshes' Version.
Private sScale As Single      ''  >> Object's Scale
                              
                              ''  Arrays:
Private Solids() As tSolid    ''  >> Array Of Solids.
Private Points() As Integer   ''  >> Temporary array of indexes.
Private tSegs() As tSeg       ''  >> Temporary array of segments.
                              
                              ''  Other Vars:
Private fID As Integer        ''  Opened File's ID.

                              ''  Rendering's Vars:
Private R As RECT             ''  A Rect.
Private SizeX As Long         ''  Display's Size.
Private SizeY As Long         ''     "        "
Private tHdc As Long          ''  Handle of DC.
Private tBit As Long          ''  Handle of BitMap.

''-----------------------------
''         Class:
''          Method &
''            Events
''------------------------------

Public Event Loading(Percentage As Single)
Public Event LoadingComplete()

Private Sub DeleteAllVars() ''  This is an initialization.
nnPoints = 0
nnSegs = 0
tnSolids = 0
tnPoints = 0
tnPolys = 0

s3dStudioV = ""
sMeshV = ""
sScale = 0
                              
Erase Solids()
Erase Points()
Erase tSegs()

End Sub

Private Function fFileLen(FileName As String) As Double
On Error GoTo notExist
fFileLen = -1
fFileLen = FileLen(FileName)
notExist:
End Function

Public Function Load3dsFile(FileName As String) As Boolean ''  This read 3ds file

Dim fPos As Long    ''  Current Read-Position On File.
Dim I As Long, fLen As Long

If fFileLen(FileName) < 0 Then GoTo Err: '' File musts exist!

Call DeleteAllVars                      ''  Delete All Vars' Values.
fID = FreeFile()                        ''  Get a free file id
Open FileName For Binary As fID         ''  Open the File
fPos = 1                                ''  Begin Of File.
fLen = LOF(fID)                         ''  File's Lenght.
RaiseEvent Loading(0)

Do While Not (EOF(fID))                      '' If the file is not ended
    Call ReadChunk(fPos)                     '' It read the next chunk.
    RaiseEvent Loading(100 * (fPos / fLen))  '' Update Loading Status.
    DoEvents: DoEvents                       '' To avoid the app. freeze.
Loop

RaiseEvent Loading(100)                 ''  Loading Complete
RaiseEvent LoadingComplete
Load3dsFile = True
Exit Function
Err:  '' Error. We exit.
Load3dsFile = False
End Function

Private Sub ReadChunk(ByRef fPos As Long) '' Read a piece.

'' The sub is recursive and the position must
'' be given BYREF!!!

Dim tChk As tChunk        ''  A chunk
Dim tInt As Integer       ''  These are temporary vars
Dim tLng As Long          ''  we use to read values
Dim tSng As Single        ''  from 3ds file.
Dim tStr As String        ''  (one for type)

Dim vInfo As Integer      ''  It tell us which segments in a triangle
                          ''  are hide.

Dim I As Integer          ''  Cicle's Var.

Get fID, fPos, tChk       ''  We read a chunk from current (fpos)
                          ''  position.

Select Case tChk.Header   ''  Every chunk give a specify information.
    Case 2                ''  3dStudio's version
        Get fID, , tLng
        tStr = CStr(tLng)
        s3dStudioV = tStr
    Case &H3D3E           ''  Meshes' Version.
        Get fID, , tLng
        tStr = CStr(tLng)
        sMeshV = tStr
    Case &H100            ''  Scale Used.
        Get fID, , tSng
        sScale = tSng
    Case &H4000           ''  The Objects!
        tStr = ReadStr()
        tLng = fPos + Len(tStr) + 7
        nnSegs = 0
        While tLng < (fPos + tChk.Length)  ''  We read all subchunks
            ReadChunk tLng
        Wend
        If tnSolids > 0 Then
            If nnSegs = 0 Then
                Solids(tnSolids).nSegs = 1
                ReDim Solids(tnSolids).Segs(1)
                ReDim Solids(tnSolids).Segs(0).Points(nnPoints)
                Solids(tnSolids).Segs(0).nPoints = nnPoints
                For I = 0 To nnPoints - 1
                    Solids(tnSolids).Segs(0).Points(I) = Points(I)
                Next
            Else
                Solids(tnSolids).nSegs = nnSegs
                ReDim Solids(tnSolids).Segs(nnSegs)
                For I = 0 To nnSegs - 1
                    ReDim Solids(tnSolids).Segs(I).Points(tSegs(I).nPoints)
                    LSet Solids(tnSolids).Segs(I) = tSegs(I)
                Next
            End If
        End If
    Case &H4110                         ''  A new Solid!!!
        tnSolids = tnSolids + 1         '' Update solids' count.
        ReDim Preserve Solids(tnSolids) '' New element on solids' array.
        Get fID, , tInt
        ReDim Solids(tnSolids).Points(tInt)
        Solids(tnSolids).nPoints = tInt
        tnPoints = nnPoints + tInt      '' Update points' count.
        For I = 0 To tInt - 1           '' Get X,Y,Z of new points.
            Get fID, , tSng             '' X
            Solids(tnSolids).Points(I).X = tSng
            Get fID, , tSng             '' Y
            Solids(tnSolids).Points(I).Y = tSng
            Get fID, , tSng             '' Z
            Solids(tnSolids).Points(I).Z = tSng
        Next
        
    Case &H4120                         '' Now we build the polys.
        Get fID, , tInt                 '' Number of polys in solid.
        tnPolys = tnPolys + tInt        '' Update polys count.
        nnPoints = CLng(tInt) * 3       '' Every triangle has 3 points.
        ReDim Points(CLng(tInt) * 3)    '' Temp. Array of points
        For I = 0 To CLng(tInt) * 3 - 1 Step 3
            Get fID, , Points(I)        '' Three points
            Get fID, , Points(I + 1)    '' for every
            Get fID, , Points(I + 2)    '' triangle.
            Get fID, , vInfo            '' Which segments are hide in
                                        '' the triangle.
        Next
        
        tLng = fPos + 8 + 8 * CLng(tInt)
        While tLng < (fPos + tChk.Length)  ''  Read Other Subchunks.
            ReadChunk tLng
        Wend
    Case &H4160 '' Local Coordinate's System
        ' I don't like it. :))
    Case &H4170 '' We skip this chunk.
        Get fID, , tInt
        tStr = CStr(tInt)
        For I = 1 To 21
            Get fID, , tSng
        Next
    Case &HA200, &HA204, &HA210, &HA220, &HA230, &H4100, &H1200, &H3000
        '' We skip also these sections...
        tLng = fPos + 6
        While tLng < (fPos + tChk.Length)  ''  Read Other Subchunks.
            ReadChunk tLng
        Wend
    Case &H4D4D, &HAFFF, &H3D3D '' Skip ... Skip ... Skip...
        tChk.Length = 6
    Case Else   '' Other Sections: Colors, Visual Effects, Keyframer... Skip!
        If tChk.Header > &HA000 And tChk.Header < &HA080 Then
            tLng = fPos + 6
        End If
End Select
fPos = fPos + tChk.Length
End Sub

Private Function ReadStr() As String '' Read an UNICODE string.
    Dim TmpChar As String
    Dim Ris As String
    TmpChar = StrConv(InputB(1, fID), vbUnicode)
    Do While TmpChar <> Chr(0)
        Ris = Ris & TmpChar
        TmpChar = StrConv(InputB(1, fID), vbUnicode)
    Loop
    ReadStr = Ris
End Function


''  XY Rotation :
''
''  X = x*cos(r) - y*sin(r)
''
''  Y = x*sin(r) - y*cos(r)

Public Sub RotateZ(Radian As Double)
Dim t1 As Double, t2 As Double
Dim tsin As Double, tCos As Double
Dim I As Long, K As Long

tsin = Sin(Radian)
tCos = Cos(Radian)

For I = 1 To tnSolids
For K = 0 To Solids(I).nPoints - 1
    t1 = Solids(I).Points(K).X
    t2 = Solids(I).Points(K).Y
    Solids(I).Points(K).X = t1 * tCos - t2 * tsin
    Solids(I).Points(K).Y = t1 * tsin + t2 * tCos
Next
Next
End Sub

Public Sub RotateY(Radian As Double)
Dim t1 As Double, t2 As Double
Dim tsin As Double, tCos As Double
Dim I As Long, K As Long

tsin = Sin(Radian)
tCos = Cos(Radian)

For I = 1 To tnSolids
For K = 0 To Solids(I).nPoints - 1
    t1 = Solids(I).Points(K).Z
    t2 = Solids(I).Points(K).X
    Solids(I).Points(K).Z = t1 * tCos - t2 * tsin
    Solids(I).Points(K).X = t1 * tsin + t2 * tCos
Next
Next
End Sub
Public Sub RotateX(Radian As Double)
Dim t1 As Double, t2 As Double
Dim tsin As Double, tCos As Double
Dim I As Long, K As Long

tsin = Sin(Radian)
tCos = Cos(Radian)

For I = 1 To tnSolids
For K = 0 To Solids(I).nPoints - 1
    t1 = Solids(I).Points(K).Z
    t2 = Solids(I).Points(K).Y
    Solids(I).Points(K).Z = t1 * tCos - t2 * tsin
    Solids(I).Points(K).Y = t1 * tsin + t2 * tCos
Next
Next
End Sub
Public Sub MoveX(qt As Double)
Dim I As Long, J As Long, K As Long
For I = 1 To tnSolids
For K = 0 To Solids(I).nPoints - 1
    Solids(I).Points(K).X = Solids(I).Points(K).X + qt
Next
Next
End Sub
Public Sub MoveY(qt As Double)
Dim I As Long, J As Long, K As Long
For I = 1 To tnSolids
For K = 0 To Solids(I).nPoints - 1
    Solids(I).Points(K).Y = Solids(I).Points(K).Y + qt
Next
Next
End Sub
Public Sub MoveZ(qt As Double)
Dim I As Long, J As Long, K As Long
For I = 1 To tnSolids
For K = 0 To Solids(I).nPoints - 1
    Solids(I).Points(K).X = Solids(I).Points(K).X * qt
    Solids(I).Points(K).Y = Solids(I).Points(K).Y * qt
    Solids(I).Points(K).Z = Solids(I).Points(K).Z * qt
Next
Next
End Sub

Public Sub Render(hdc As Long)
Dim cOld As POINTAPI
Dim pPoly() As POINTAPI
Dim I As Long, J As Long, K As Long

'' Create the background  >>>

FillRect tHdc, R, 255 'GetSysColorBrush(5)

For I = 1 To tnSolids
    For K = 0 To Solids(I).nSegs - 1
        ReDim pPoly(Solids(I).Segs(K).nPoints - 1)
        
        ''  Create a chain of point >>>
        
        For J = 0 To Solids(I).Segs(K).nPoints - 1
            pPoly(J).X = Solids(I).Points(Solids(I).Segs(K).Points(J)).X + SizeX / 2
            pPoly(J).Y = Solids(I).Points(Solids(I).Segs(K).Points(J)).Y + SizeY / 2
        Next
        
        ''  Move the pen to the first point of the chain >>>
        
        MoveToEx tHdc, pPoly(0).X, pPoly(0).Y, cOld
        
        ''  Draw the chain of points >>>
        
        For J = 0 To UBound(pPoly) - 1
            LineTo tHdc, pPoly(J + 1).X, pPoly(J + 1).Y
        Next
        
    Next
Next

''  Copy on Final DC >>>
BitBlt hdc, 0, 0, SizeX, SizeY, tHdc, 0, 0, vbSrcCopy

''  *****************************************************
''  IF WE WORK ON A BUFFER AND THEN COPY IT ON PICTUREBOX
''  WE AVOID THE FLIP EFFECT
''  *****************************************************

End Sub

Public Sub CreateBuffer(Pic As PictureBox)
Pic.ScaleMode = 3                      ''  Windows' API use Pixel.
tHdc = CreateCompatibleDC(Pic.hdc)     ''  Create a DC & a BitMap
tBit = CreateCompatibleBitmap(tHdc, Pic.ScaleWidth, Pic.ScaleHeight)
SizeX = Pic.ScaleWidth                 ''  Size of display
SizeY = Pic.ScaleHeight                ''  ...
DeleteObject SelectObject(tHdc, tBit)  ''  Set BitMap
SetRect R, 0, 0, SizeX, SizeY          ''  Set Rect.
End Sub

Public Sub DeleteBuffer()
''  Free Memory Space.
DeleteDC tHdc
DeleteObject tBit
End Sub

Public Property Get SolidsCount()
    SolidsCount = tnSolids
End Property

Public Property Get PointsCount()
    PointsCount = tnPoints
End Property

Public Property Get PolysCount()
    PolysCount = tnPolys
End Property


