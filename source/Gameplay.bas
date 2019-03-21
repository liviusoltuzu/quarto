Attribute VB_Name = "Gameplay"
Option Explicit

Public StilDeJoc As Byte '1,2,3=beginner, intermediate, advanced
'Public Inceput As Boolean 'spune dk a fost facuta o prima mutare (calc sau utiliz)
Public SingleOption(1 To 4) As Boolean 'daca este beginner are posib sa-si selecteze doar
                                        'anumite caracteristici
Public Ka As Byte 'cate posibilitati de a pune piesa pe tabla
Public Ilose As Boolean 'cand pierde calculatorul
Public Youlose As Boolean 'cand pierde userul
Public Remiza As Boolean

Public Victs As Integer 'modificat in ai la calccomb pentru a afisa la ilose=1 sau youlose=1
Type VictCateg
    ACuloare As Byte
    BMarime As Byte
    CForma As Byte
    DSoliditate As Byte
End Type
'1-linia1 2-linia2 3-linia3 4-linia4 ...... 5-col1 6-col2 7-col3 8-col4
'....9diagprinc 10diagsec
'11-1patra 12-2p 13-3p 14-4p 15-5p 16-6p 17-7p 18-8p 19-9p
Public ArrayVicts(1 To 19) As VictCateg

Public CineMuta As Boolean 'true=calculator false=user. la inceput=new game
Public CineMutaAcum As Boolean 'true=calculator false=user
Public WhoStarts As Byte '0=alternative 1=player 2=computer

'tempuri
Public WhoStartsTemp As Byte 'pentru options ca sa vada dk s-a schimbat ceva
Public StilDeJocTemp As Byte
Public SingleOptionTemp(1 To 4) As Boolean

Public PiesaForU As Byte 'este piesa pe care o da calcul userului
Public PiesaForC As Byte 'este piesa pe care o da userul calcului
Public TbSaDeaPiesa As Boolean 'daca userul tb sa aleaga piesa
Public AvPieces(1 To 16) As Byte 'piesele disponibile din Piese
Public ShowPiece(1 To 16) As Boolean 'vector care permite afisarea doar a unor piese
Public MenuPieces(1 To 13) As Boolean
Public AllPieces As Byte 'toate piesele disponibile
Public ktLightA As Byte, ktDarkA As Byte, ktTallA As Byte, ktShortA As Byte, ktHollowA As Byte
Public ktLight As Byte, ktDark As Byte, ktTall As Byte, ktShort As Byte, ktHollow As Byte
Public ktSolidA As Byte, ktRoundA As Byte, ktSquareA As Byte
Public ktSolid As Byte, ktRound As Byte, ktSquare As Byte 'cele cu a=per total/cele fara intra la ktpiese


Public Sub NewGame()
On Error Resume Next
Dim i As Byte
If LogF <> 0 Then Print #LogF, GetTime & "New game"

Select Case WhoStarts
Case 0
    If CineMuta = False Then
        CineMuta = True
    Else
        CineMuta = False
    End If
Case 1
    CineMuta = False
Case 2
    CineMuta = True
End Select
CineMutaAcum = CineMuta

CreeazaArrayPiese
CreateRandomArrays
ResetTabla
ResetArrayComb
'ReadSett

Ilose = False
Youlose = False
Remiza = False
Dim j As Byte
For i = 1 To 4: For j = 1 To 4: TblDblClk(i, j) = 0: Next: Next
'for i=1 to 4 lsttblclk()


XYZCoord(20).coord = 1
XYZCoord(21).coord = 1
For i = 1 To 16
    XYZCoord(i).coord = CDbl(-0.5)
    XYZCoord(i).bool = True
Next

For i = 1 To 12
    MenuPieces(i) = True
Next

'reset lights
aflLightPosition(0) = 0#
aflLightPosition(1) = 1#
aflLightPosition(2) = 1#
aflLightPosition(3) = 1#
  
glLightfv ltLight1, lpmPosition, aflLightPosition(0)
'end lights

Dim OPiesa As Piesa
OPiesa.ACuloare = 2
OPiesa.BMarime = 2
OPiesa.CForma = 2
OPiesa.DSoliditate = 2
OPiesa.ESett = 0
If CineMutaAcum Then
    PiesaForU = DaPiesa(OPiesa)
    CineMutaAcum = False
    TbSaDeaPiesa = False
Else
    TbSaDeaPiesa = True
    KeyControl = 1
End If
KTpiese = CreateAvPieces
PosTabla = 1
PosPickList = 1
If LogF <> 0 Then Print #LogF, GetTime & "Computer starts: " & CineMuta
If CineMuta Then
    If LogF <> 0 Then Print #LogF, GetTime & "computer gives " & PiesaForU
End If


End Sub


Public Sub MutaCalc()
Dim OPiesa As Piesa, i As Byte

PunePiesa Piese(PiesaForC)
If Ka <> 0 Then
    Dim Ka2 As Byte
    Ka2 = 1
    For i = 2 To Ka
        If VerifCombLaFel(ArrayCombinatiiPunePiesa(1), ArrayCombinatiiPunePiesa(i)) Then
            Ka2 = Ka2 + 1
        Else
            Exit For
        End If
    Next
    Randomize
    Ka2 = Int(Rnd * Ka2 + 1)
    
    Piese(PiesaForC).ESett = 1
    'While DelayPutGivenP > Timer
    '    DoEvents
    'Wend
    Tabla(ArrayCombinatiiPunePiesa(Ka2).PozX, ArrayCombinatiiPunePiesa(Ka2).PozY) = Piese(PiesaForC)
    If LogF <> 0 Then Print #LogF, GetTime & "computer on " & ArrayCombinatiiPunePiesa(Ka2).PozX & ArrayCombinatiiPunePiesa(Ka2).PozY
    
    
    If ArrayCombinatiiPunePiesa(Ka2).Victory = 0 Then
        OPiesa = ArrayCombinatiiPunePiesa(Ka2).CePiesaTreb
        If OPiesa.ACuloare = 2 Or OPiesa.BMarime = 2 Or OPiesa.CForma = 2 Or OPiesa.DSoliditate = 2 Then
            PiesaForU = DaPiesa(OPiesa)
            If PiesaForU = 0 Then
                Remiza = True
                If LogF <> 0 Then Print #LogF, GetTime & "Draw!"
                Exit Sub
            End If
            Piese(PiesaForU).ESett = 1
            If LogF <> 0 Then Print #LogF, GetTime & "computer gives " & PiesaForU
            KTpiese = CreateAvPieces
        Else
            For i = 1 To 16
                If OPiesa.ACuloare = Piese(i).ACuloare And OPiesa.BMarime = Piese(i).BMarime And _
                OPiesa.CForma = Piese(i).CForma And OPiesa.DSoliditate = Piese(i).DSoliditate Then
                    Exit For
                End If
            Next
            PiesaForU = i
            Piese(PiesaForU).ESett = 1
            If LogF <> 0 Then Print #LogF, GetTime & "computer gives " & PiesaForU
            KTpiese = CreateAvPieces
        End If
    Else
        Youlose = True
        If LogF <> 0 Then Print #LogF, GetTime & "You lose!"
        CalculeazaComb ArrayCombinatiiPunePiesa(Ka2).PozX, ArrayCombinatiiPunePiesa(Ka2).PozY, True
    End If
Else
    Dim Care As Byte 'care pozitie va fi aleasa
    If kAny <> 0 Then
    Care = 1
        For i = 2 To kAny
            If ArrayCPPTemp(i).LiniiDe3 < ArrayCPPTemp(Care).LiniiDe3 Then
                Care = i
            End If
        Next
    Care = Int(Rnd * Care + 1)
    Tabla(ArrayCPPTemp(Care).PozX, ArrayCPPTemp(Care).PozY) = Piese(PiesaForC)
    If LogF <> 0 Then Print #LogF, GetTime & "computer on " & ArrayCPPTemp(Care).PozX & ArrayCPPTemp(Care).PozY
    OPiesa.ACuloare = 2
    OPiesa.BMarime = 2
    OPiesa.CForma = 2
    OPiesa.DSoliditate = 2
    PiesaForU = DaPiesa(OPiesa)
    If LogF <> 0 Then Print #LogF, GetTime & "computer gives " & PiesaForU
    KTpiese = CreateAvPieces
    End If
End If
'KeyControl = 2

End Sub

Public Function CreateAvPieces() As Byte 'creeaza un vector cu piesele disponibile available
Dim l As Byte, k As Byte, b As Boolean
k = 0
AllPieces = 0
 ktLightA = 0: ktDarkA = 0: ktTallA = 0: ktShortA = 0: ktHollowA = 0:
 ktLight = 0: ktDark = 0: ktTall = 0: ktShort = 0: ktHollow = 0:
 ktSolidA = 0: ktRoundA = 0: ktSquareA = 0:
 ktSolid = 0: ktRound = 0: ktSquare = 0:

For l = 1 To 16
    b = True
    If Piese(l).ACuloare = 0 And Not (MenuPieces(1) Or MenuPieces(2)) Then b = False
    If Piese(l).ACuloare = 1 And Not (MenuPieces(1) Or MenuPieces(3)) Then b = False
    
    If Piese(l).BMarime = 0 And Not (MenuPieces(4) Or MenuPieces(5)) Then b = False
    If Piese(l).BMarime = 1 And Not (MenuPieces(4) Or MenuPieces(6)) Then b = False
    
    If Piese(l).CForma = 0 And Not (MenuPieces(7) Or MenuPieces(8)) Then b = False
    If Piese(l).CForma = 1 And Not (MenuPieces(7) Or MenuPieces(9)) Then b = False
    
    If Piese(l).DSoliditate = 0 And Not (MenuPieces(10) Or MenuPieces(11)) Then b = False
    If Piese(l).DSoliditate = 1 And Not (MenuPieces(10) Or MenuPieces(12)) Then b = False
    
    If Piese(l).ESett = 0 Then
        AllPieces = AllPieces + 1
        increaseKT l, True
        If b Then
            k = k + 1
            increaseKT l, False
            AvPieces(k) = l
        End If
    End If
Next

CreateAvPieces = k
End Function

Private Sub increaseKT(nrps As Byte, tot As Boolean)
If tot Then
    If Piese(nrps).ACuloare = 0 Then ktLightA = ktLightA + 1
    If Piese(nrps).ACuloare = 1 Then ktDarkA = ktDarkA + 1

    If Piese(nrps).BMarime = 0 Then ktShortA = ktShortA + 1
    If Piese(nrps).BMarime = 1 Then ktTallA = ktTallA + 1

    If Piese(nrps).CForma = 0 Then ktRoundA = ktRoundA + 1
    If Piese(nrps).CForma = 1 Then ktSquareA = ktSquareA + 1

    If Piese(nrps).DSoliditate = 0 Then ktHollowA = ktHollowA + 1
    If Piese(nrps).DSoliditate = 1 Then ktSolidA = ktSolidA + 1
Else
    If Piese(nrps).ACuloare = 0 Then ktLight = ktLight + 1
    If Piese(nrps).ACuloare = 1 Then ktDark = ktDark + 1

    If Piese(nrps).BMarime = 0 Then ktShort = ktShort + 1
    If Piese(nrps).BMarime = 1 Then ktTall = ktTall + 1

    If Piese(nrps).CForma = 0 Then ktRound = ktRound + 1
    If Piese(nrps).CForma = 1 Then ktSquare = ktSquare + 1

    If Piese(nrps).DSoliditate = 0 Then ktHollow = ktHollow + 1
    If Piese(nrps).DSoliditate = 1 Then ktSolid = ktSolid + 1
End If
End Sub






