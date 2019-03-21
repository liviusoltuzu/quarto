Attribute VB_Name = "GLLists"
Option Explicit
Type ThreeCoords
    vX As Double
    vY As Double
    vZ As Double
End Type

Type TwoCoords
    vX As Double
    vY As Double
End Type

Type TDVertex
    V As Long
    t As Long
    N As Long
End Type

Type TDFaces
    Vertex(1 To 3) As TDVertex
End Type

Type TDObj
    ObjName As String
    NumF As Long
    Faces() As TDFaces
End Type
    
Type TDObject
    objs() As TDObj
    NumV As Long
    NumT As Long
    NumN As Long
    Verticies() As ThreeCoords
    TCoords() As TwoCoords 'texture
    Normals() As ThreeCoords
End Type
    
Global Objects As TDObject
Type TDGLList
    ObjName As String
    Idx As Byte
End Type
Public Lists(1 To 26) As TDGLList


Public Sub ReadObj(OBJPath As String)
On Error GoTo err:
Dim a() As String, b() As String, C() As String
Dim F As String
'Dim MinusT As Long, MinusN As Long, MinusV As Long
Dim i As Long, j As Long
Dim CurObj As Byte
Dim FileNumber As Integer
Dim Ds As String * 1 'decimal separator
'Dim tz As Long

If LogF <> 0 Then Print #LogF, GetTime & "reading files..."
ZStart.lblDoing = "Reading files..."
ZStart.lblDoing.Refresh
'DoEvents
FileNumber = FreeFile
Open OBJPath For Input As #FreeFile
Dim Lseek As Long 'line seek
Lseek = 0
While Not EOF(FileNumber)
    If Lseek Mod 751 = 0 Then DoEvents
    Lseek = Lseek + 1
    ReDim Preserve a(Lseek)
    Line Input #FileNumber, a(Lseek)
Wend
Close #FileNumber
'Kill OBJPath
CurObj = 0
If LogF <> 0 Then Print #LogF, GetTime & "done."
If LogF <> 0 Then Print #LogF, GetTime & "loading objects..."
ZStart.lblDoing = "Loading objects..."
ZStart.lblDoing.Refresh

Ds = Format$(0, ".")
With Objects
For i = 1 To UBound(a) - 1
    If i Mod 751 = 0 Then DoEvents
    b = Split(a(i), " ")
    Select Case LCase(b(0))
    Case "g"
        CurObj = CurObj + 1
        ReDim Preserve Objects.objs(CurObj)
        Objects.objs(CurObj).ObjName = b(1)
    Case "v"
        .NumV = .NumV + 1
        ReDim Preserve .Verticies(.NumV)
        .Verticies(.NumV).vX = Replace(b(1), ".", Ds)
        .Verticies(.NumV).vY = Replace(b(2), ".", Ds)
        .Verticies(.NumV).vZ = Replace(b(3), ".", Ds)
    Case "vt"
        .NumT = .NumT + 1
        ReDim Preserve .TCoords(.NumT)
        .TCoords(.NumT).vX = Replace(b(1), ".", Ds)
        .TCoords(.NumT).vY = Replace(b(2), ".", Ds)
    Case "vn"
        .NumN = .NumN + 1
        ReDim Preserve .Normals(.NumN)
        .Normals(.NumN).vX = Replace(b(1), ".", Ds)
        .Normals(.NumN).vY = Replace(b(2), ".", Ds)
        .Normals(.NumN).vZ = Replace(b(3), ".", Ds)
    Case "f"
        With Objects.objs(CurObj)
        .NumF = .NumF + 1
        ReDim Preserve .Faces(.NumF)
        For j = 1 To 3
            C = Split(b(j + 1), "/")
            .Faces(.NumF).Vertex(j).V = C(0)
            .Faces(.NumF).Vertex(j).t = C(1)
            .Faces(.NumF).Vertex(j).N = C(2)
        Next
        End With
End Select
Next
End With

Erase a
Erase b
Erase C
F = ""
If LogF <> 0 Then Print #LogF, GetTime & "done."
err:
If err.Number <> 0 Then
    ErrPrg = True
    ErrNumb = 101
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    End
End If
End Sub

Public Sub CreateGLLists()
Dim i As Byte, j As Byte
With Objects
For i = 1 To 8
    DoEvents
    j = Mid(.objs(i).ObjName, 2, 1)
    Lists(j).ObjName = .objs(i).ObjName
    Lists(j).Idx = j
    glNewList j, lstCompile
    CreateList i
    glEndList
Next
For i = 9 To 10
    DoEvents
    Lists(i).Idx = i
    Lists(i).ObjName = .objs(i).ObjName
    glNewList i, lstCompile
    CreateList i
    glEndList
Next
For i = 11 To 26
    DoEvents
    j = (Mid(.objs(i).ObjName, 2, 1) - 1) * 4 + Mid(.objs(i).ObjName, 3, 1)
    Lists(10 + j).ObjName = .objs(i).ObjName
    Lists(10 + j).Idx = i
    glNewList 10 + j, lstCompile
    CreateList i
    glEndList
Next
End With
Rem 27 este picklist
'GLCreateShape HWGivenPiece, HWGivenPiece, 28, False, , , True 'viewport dreapta-sus/dreapta-jos
GLCreateShape HWPiece, HWPiece, 29 'viewport sus cu piese
GLCreateShape WidthPickListButton, HeightPickListButton, 30, True, 2 'buton
'31 este menu
GLCreateTriangle WidthPLBTriangle, HeightPLBTriangle, 32, False
GLCreateTriangle WidthPLBTriangle, HeightPLBTriangle, 33, True
GLCreateShape HWGivenPiece, HWGivenPiece, 34 'patrat giveNpiece

GLCreateShape WidthMenuTitle, HeightMenuTitle, 35, False 'titlu meniu
GLCreateShape WidthMenuTitle, HeightMenuTitle, 35, False 'titlu meniu rolled
GLCreateShape WidthRolledMenu, HeightRolledMenu, 36, False 'element din meniu

GLCreateShape WidthOptions, HeightOptions, 37, False   'options window
GLCreateShape 32, 32, 38, False
'option box
GLCreateShape 12, 12, 39, True, 6, 3, False 'cerc mare negru
GLCreateShape 8, 8, 40, True, 4, 3, False 'cerc mare alb
GLCreateShape 4, 4, 41, True, 2, 2, False 'bila =bifat
'check box
GLCreateShape 13, 13, 42, False 'patrat mare negru
GLCreateShape 9, 9, 43, False 'patrat mare alb
GLCreateShape 5, 5, 44, 0 'patrat mic =bifat

GLCreateShape WidthButton, HeightButton, 45, True
'mesaj
GLCreateShape WidthMessage, HeightMessage, 46, True
GLCreateShape WidthMessage - 7, HeightMessage - 7, 47, False, , , True

BuildFont Stage2.Stage
BuildFont Stage2.Stage, 1
BuildFont Stage2.Stage, 2


End Sub

Private Sub CreateList(ObjNmb As Byte, Optional CrAs As Byte = 4)
Dim i As Long, j As Byte
Dim vX As Double, vY As Double, vZ As Double
Dim vNx As Double, vNy As Double, vNz As Double
Dim vTx As Double, vTy As Double

glBegin CrAs
With Objects
For i = 1 To .objs(ObjNmb).NumF
    For j = 1 To 3
        vX = .Verticies(.objs(ObjNmb).Faces(i).Vertex(j).V).vX
        vY = .Verticies(.objs(ObjNmb).Faces(i).Vertex(j).V).vY
        vZ = .Verticies(.objs(ObjNmb).Faces(i).Vertex(j).V).vZ
        vTx = .TCoords(.objs(ObjNmb).Faces(i).Vertex(j).t).vX
        vTy = .TCoords(.objs(ObjNmb).Faces(i).Vertex(j).t).vY
        vNx = .Normals(.objs(ObjNmb).Faces(i).Vertex(j).N).vX
        vNy = .Normals(.objs(ObjNmb).Faces(i).Vertex(j).N).vY
        vNz = .Normals(.objs(ObjNmb).Faces(i).Vertex(j).N).vZ
        glNormal3f vNx, vNy, vNz
        glTexCoord2f vTx, vTy
        glVertex3f vX / ZoomRate, vY / ZoomRate, vZ / ZoomRate
    Next
Next
End With
glEnd
End Sub
