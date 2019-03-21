Attribute VB_Name = "AI"
Option Explicit
Public Type Piesa
    '2=nul
    ACuloare As Byte '0=alb 1=negru
    BMarime As Byte '0=mic 1=mare
    CForma As Byte '0=rotund 1=patrat
    DSoliditate As Byte '0=gaurit 1=plat
    ESett As Byte 'in calcomb: dk arepiesa=bool, in piese(array):dk a fost sau
                    'nu folosita, in tabla(array) dk e liber sau nu
                    
                    'piese(array): 0=libera 1=folosita
                    'tabla(array2): 0=poz libera     1=poz ocupata
End Type

Public Type Combinatie
    LiniiDe2 As Byte 'cate linii de 2 sunt pe tabla
    LiniiDe3 As Byte 'cate linii de 3 sunt pe tabla
    Victory As Byte 'cate victorii "-"
    NoComb As Boolean 'in caz ca nu exista nici linii, nici victorii...
    Block As Byte 'cate blocari de posibile combinatii
    VictOnCuloare As Byte 'dk victory spune la ce categ
    VictOnMarime As Byte '""
    VictOnForma As Byte '""
    VictOnSoliditate As Byte '""
    PozX As Byte 'unde pe tabla avem piesa = x
    PozY As Byte '=y
    ArePiesa As Boolean 'dk pentru combinatia asta are piesa sa dea!!! in caz
                        'de linie de 3
    CePiesaTreb As Piesa 'ce caracteristici are piesa pe care tb sa o dea!!!
End Type

Private ArrayA(1 To 9) As Byte 'pentru random-ul pe care il face la pune piesa=    'beginner
Private ArrayB(1 To 9) As Byte '=ce mutare face: create3 (face o lin de 3),create2   'intermediate
Private ArrayC(1 To 9) As Byte 'nocombination (mai rar...), block                       'advanced
                                        'create3=1; create2=2; block=3; nocombination=4
                                        '(victorie nu intra aici)
Private Noua(1 To 3) As Byte 'cate elemente mai sunt in vectorii de random


Public kAny As Integer 'cate combinatii sunt, chiar daca pierde sau nu
Public Tabla(1 To 4, 1 To 4) As Piesa 'tabla
Public Piese(1 To 16) As Piesa 'vector cu tt piesele, folosite sau nu
Public PiesaCurenta As Piesa 'piesa pe care tb sa o puna
Public ArrayCombinatiiPunePiesa(1 To 16) As Combinatie 'informatiile despre fiecare mutare
                                                    'posibila (vezi Type Combinatie)
Public ArrayCPPTemp(1 To 16) As Combinatie  'vezi un rand mai sus: pentru cazul cand pierde

Public Sub PunePiesa(Ce As Piesa)
Dim k As Byte 'cate poz libere sunt pe tabla = counter din ArrayCombinatiiPunePiesa
Dim X As Byte
Dim Y As Byte
Dim i As Byte
Dim j As Byte
Dim Vict As Boolean 'cand parcurge array retine dk a fost gasita o victorie
Dim TempBool As Boolean 'true dk a fost gasit in arraycombpunepiesa linde3, de2,blk sau nc
Dim Rn As Byte 'ce mutare va face: create3,2,blk,nc luata cu rnd-ul
PiesaCurenta = Ce
Randomize
'Inceput = True '!!!!!!!!!!!!!!!!!!!    'doar in faza de teste: dupa aceea tb stersa linia asta
'If Not Inceput Then
'    X = Int((4 * Rnd) + 1)              'dk e prima mutare a calc, o pune la intamplare
'    Y = Int((4 * Rnd) + 1)
'    ArrayCombinatiiPunePiesa(1).PozX = X
'    ArrayCombinatiiPunePiesa(1).PozY = Y
'    'Tabla(x, y) = PiesaCurenta
'    Inceput = True
'Else
    'kAny = 0
    k = 0
    For i = 1 To 4
        For j = 1 To 4
            If Tabla(i, j).ESett = 0 Then   'dk e libere poz calc combinatii pe ea
                Tabla(i, j) = PiesaCurenta  'temporar, vede cum ar fi cu piesa in poz aceasta (A)
                k = k + 1
                ArrayCombinatiiPunePiesa(k) = CalculeazaComb(i, j)
                Tabla(i, j) = ResetPiesa(Tabla(i, j))   'sterge testul (A2)
            End If
        Next
    Next
'efectiv pune=adica gaseste cea mai buna mutare din arraycombpunepiesa...
    CreateRandomArrays
    For i = 1 To k
        If ArrayCombinatiiPunePiesa(i).Victory <> 0 Then
            Vict = True
            Exit For
        End If
    Next
    If Vict Then
        i = 1
        While i <= k
            If ArrayCombinatiiPunePiesa(i).Victory = 0 Then
                DeleteEntry 0, i, k
                k = k - 1
            Else: i = i + 1
            End If
        Wend
        SortArrayComb 0, k
    Else
        i = 1
        kAny = 0
        While i <= k
            If Not ArrayCombinatiiPunePiesa(i).ArePiesa Then
                kAny = kAny + 1
                ArrayCPPTemp(kAny) = ArrayCombinatiiPunePiesa(i)
                DeleteEntry 0, i, k
                k = k - 1
            Else: i = i + 1
            End If
        Wend
        
        'linii de 3......
        TempBool = False
        For i = 1 To k
            If ArrayCombinatiiPunePiesa(i).LiniiDe3 <> 0 Then
                TempBool = True
                Exit For
            End If
        Next
        If Not TempBool Then
            i = 1
            While i <= Noua(1)
                If ArrayA(i) = 1 Then
                DeleteEntry 1, i, Noua(1): Noua(1) = Noua(1) - 1
                Else: i = i + 1
                End If
            Wend
            i = 1
            While i <= Noua(2)
                If ArrayB(i) = 1 Then
                DeleteEntry 2, i, Noua(2): Noua(2) = Noua(2) - 1
                Else: i = i + 1
                End If
            Wend
            i = 1
            While i < Noua(3)
                If ArrayC(i) = 1 Then
                DeleteEntry 3, i, Noua(3): Noua(3) = Noua(3) - 1
                Else: i = i + 1
                End If
            Wend
        End If
                
        'linii de 2......
        TempBool = False
        For i = 1 To k
            If ArrayCombinatiiPunePiesa(i).LiniiDe2 <> 0 Then
                TempBool = True
                Exit For
            End If
        Next
        If Not TempBool Then
            i = 1
            While i <= Noua(1)
                If ArrayA(i) = 2 Then
                DeleteEntry 1, i, Noua(1): Noua(1) = Noua(1) - 1
                Else: i = i + 1
                End If
            Wend
            i = 1
            While i <= Noua(2)
                If ArrayB(i) = 2 Then
                DeleteEntry 2, i, Noua(2): Noua(2) = Noua(2) - 1
                Else: i = i + 1
                End If
            Wend
            i = 1
            While i <= Noua(3)
                If ArrayC(i) = 2 Then
                DeleteEntry 3, i, Noua(3): Noua(3) = Noua(3) - 1
                Else: i = i + 1
                End If
            Wend
        End If
        
        'blk......
        TempBool = False
        For i = 1 To k
            If ArrayCombinatiiPunePiesa(i).Block <> 0 Then
                TempBool = True
                Exit For
            End If
        Next
        If Not TempBool Then
            i = 1
            While i <= Noua(1)
                If ArrayA(i) = 3 Then
                DeleteEntry 1, i, Noua(1): Noua(1) = Noua(1) - 1
                Else: i = i + 1
                End If
            Wend
            i = 1
            While i <= Noua(2)
                If ArrayB(i) = 3 Then
                DeleteEntry 2, i, Noua(2): Noua(2) = Noua(2) - 1
                Else: i = i + 1
                End If
            Wend
            i = 1
            While i <= Noua(3)
                If ArrayC(i) = 3 Then
                DeleteEntry 3, i, Noua(3): Noua(3) = Noua(3) - 1
                Else: i = i + 1
                End If
            Wend
        End If
        
        'nc......
        TempBool = False
        For i = 1 To k
            If ArrayCombinatiiPunePiesa(i).NoComb = True Then
                TempBool = True
                Exit For
            End If
        Next
        If Not TempBool Then
            i = 1
            While i <= Noua(1)
                If ArrayA(i) = 4 Then
                DeleteEntry 1, i, Noua(1): Noua(1) = Noua(1) - 1
                Else: i = i + 1
                End If
            Wend
            'nu contine 4(nc) arrayb-ul si arrayc-ul
        End If
        'face random
        Rn = Int((Rnd * Noua(StilDeJoc)) + 1)
        Ka = k
        If k <> 0 Then
            SortArrayComb Rn, k
        End If
    End If
              
              
    

End Sub

Public Function DaPiesa(Ce As Piesa) As Byte
'explicatie in verifexistapiesa
Dim i As Byte
Dim Temp As Piesa, k As Byte
'!!!!!!!!!!pentru complexitate declaratia urmatoare cu '
'Dim ArrayDaPiesa(1 To 16) As Combinatie
Dim CePiesePoateDa(1 To 16) As Byte
k = 0
For i = 1 To 16
    If Piese(i).ESett = 0 Then
        Temp = Piese(i)
        If Ce.ACuloare = 2 Then Temp.ACuloare = 2
        If Ce.BMarime = 2 Then Temp.BMarime = 2
        If Ce.CForma = 2 Then Temp.CForma = 2
        If Ce.DSoliditate = 2 Then Temp.DSoliditate = 2
        
        If Temp.ACuloare = Ce.ACuloare And Temp.BMarime = Ce.BMarime And _
        Temp.CForma = Ce.CForma And Temp.DSoliditate = Ce.DSoliditate Then
            k = k + 1
            CePiesePoateDa(k) = i
            'CePiesePoateDa(k).ESett = 1
        End If
        
        '!!!!!!!ptr complexitate
        'If Temp.ACuloare = Ce.ACuloare And Temp.BMarime = Ce.BMarime And _
        'Temp.CForma = Ce.CForma And Temp.DSoliditate = Ce.DSoliditate Then
        '    PunePiesa Piese(i)
        '    k = k + 1
        '    ArrayDaPiesa(k) = ArrayCombinatiiPunePiesa(1)
        'End If
        
    End If
Next
Randomize
If k <> 0 Then
i = CePiesePoateDa(Int((Rnd * k) + 1))
Piese(i).ESett = 1
DaPiesa = i
Else
DaPiesa = 0
End If


End Function
Private Sub SortArrayComb(Rn As Byte, k As Byte)
'0=victory!!! de tratat cu vom,voc,vof,vos.......done
Dim i As Byte, j As Byte
Dim VictOnCateI As Byte, VictOnCateJ As Byte
If Rn <> 0 Then
Select Case StilDeJoc
Case 1
    Rn = ArrayA(Rn)
Case 2
    Rn = ArrayB(Rn)
Case 3
    Rn = ArrayC(Rn)
End Select
End If

Select Case Rn
Case 0
    For i = 1 To k - 1
        For j = i + 1 To k
            If ArrayCombinatiiPunePiesa(i).Victory < ArrayCombinatiiPunePiesa(j).Victory Then
                Schimba i, j
            ElseIf ArrayCombinatiiPunePiesa(i).Victory = ArrayCombinatiiPunePiesa(j).Victory Then
                VictOnCateI = 0
                If ArrayCombinatiiPunePiesa(i).VictOnCuloare <> 2 Then VictOnCateI = VictOnCateI + 1
                If ArrayCombinatiiPunePiesa(i).VictOnForma <> 2 Then VictOnCateI = VictOnCateI + 1
                If ArrayCombinatiiPunePiesa(i).VictOnMarime <> 2 Then VictOnCateI = VictOnCateI + 1
                If ArrayCombinatiiPunePiesa(i).VictOnSoliditate <> 2 Then VictOnCateI = VictOnCateI + 1
                
                VictOnCateJ = 0
                If ArrayCombinatiiPunePiesa(j).VictOnCuloare <> 2 Then VictOnCateJ = VictOnCateJ + 1
                If ArrayCombinatiiPunePiesa(j).VictOnForma <> 2 Then VictOnCateJ = VictOnCateJ + 1
                If ArrayCombinatiiPunePiesa(j).VictOnMarime <> 2 Then VictOnCateJ = VictOnCateJ + 1
                If ArrayCombinatiiPunePiesa(j).VictOnSoliditate <> 2 Then VictOnCateJ = VictOnCateJ + 1
                
                If VictOnCateI < VictOnCateJ Then
                    Schimba i, j
                End If
            End If
        Next
    Next
Case 1
    For i = 1 To k - 1
        For j = i + 1 To k
            If ArrayCombinatiiPunePiesa(i).LiniiDe3 < ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                Schimba i, j
            ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe3 = ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                If ArrayCombinatiiPunePiesa(i).LiniiDe2 < ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                    Schimba i, j
                ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe2 = ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                    If ArrayCombinatiiPunePiesa(i).Block < ArrayCombinatiiPunePiesa(j).Block Then
                        Schimba i, j
                    ElseIf ArrayCombinatiiPunePiesa(i).Block = ArrayCombinatiiPunePiesa(j).Block Then
                        If ArrayCombinatiiPunePiesa(i).NoComb > ArrayCombinatiiPunePiesa(j).NoComb Then
                            Schimba i, j
                        End If
                    End If
                End If
            End If
        Next
    Next
Case 2
    For i = 1 To k - 1
        For j = i + 1 To k
            If ArrayCombinatiiPunePiesa(i).LiniiDe2 < ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                Schimba i, j
            ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe2 = ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                If ArrayCombinatiiPunePiesa(i).LiniiDe3 < ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                    Schimba i, j
                ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe3 = ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                    If ArrayCombinatiiPunePiesa(i).Block < ArrayCombinatiiPunePiesa(j).Block Then
                        Schimba i, j
                    ElseIf ArrayCombinatiiPunePiesa(i).Block = ArrayCombinatiiPunePiesa(j).Block Then
                        If ArrayCombinatiiPunePiesa(i).NoComb > ArrayCombinatiiPunePiesa(j).NoComb Then
                            Schimba i, j
                        End If
                    End If
                End If
            End If
        Next
    Next
Case 3
    For i = 1 To k - 1
        For j = i + 1 To k
            If ArrayCombinatiiPunePiesa(i).Block < ArrayCombinatiiPunePiesa(j).Block Then
                Schimba i, j
            ElseIf ArrayCombinatiiPunePiesa(i).Block = ArrayCombinatiiPunePiesa(j).Block Then
                If ArrayCombinatiiPunePiesa(i).LiniiDe3 < ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                    Schimba i, j
                ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe3 = ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                    If ArrayCombinatiiPunePiesa(i).LiniiDe2 < ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                        Schimba i, j
                    ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe2 = ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                        If ArrayCombinatiiPunePiesa(i).NoComb > ArrayCombinatiiPunePiesa(j).NoComb Then
                            Schimba i, j
                        End If
                    End If
                End If
            End If
        Next
    Next
Case 4
    For i = 1 To k - 1
        For j = i + 1 To k
            If ArrayCombinatiiPunePiesa(i).NoComb > ArrayCombinatiiPunePiesa(i).NoComb Then
                Schimba i, j
            ElseIf ArrayCombinatiiPunePiesa(i).NoComb = ArrayCombinatiiPunePiesa(j).NoComb Then
                If ArrayCombinatiiPunePiesa(i).LiniiDe3 < ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                    Schimba i, j
                ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe3 = ArrayCombinatiiPunePiesa(j).LiniiDe3 Then
                    If ArrayCombinatiiPunePiesa(i).LiniiDe2 < ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                        Schimba i, j
                    ElseIf ArrayCombinatiiPunePiesa(i).LiniiDe2 = ArrayCombinatiiPunePiesa(j).LiniiDe2 Then
                        If ArrayCombinatiiPunePiesa(i).Block < ArrayCombinatiiPunePiesa(j).Block Then
                            Schimba i, j
                        End If
                    End If
                End If
            End If
        Next
    Next
End Select
            
            
                
End Sub
Private Sub Schimba(Cei As Byte, Cej As Byte)
Dim aux As Combinatie
aux = ArrayCombinatiiPunePiesa(Cei)
ArrayCombinatiiPunePiesa(Cei) = ArrayCombinatiiPunePiesa(Cej)
ArrayCombinatiiPunePiesa(Cej) = aux
End Sub

Public Sub CreateRandomArrays()
'1=create3
'2=create2
'3=blk
'4=nocombination
Noua(1) = 9
Noua(2) = 9
Noua(3) = 9
ArrayA(1) = 1: ArrayA(2) = 1: ArrayA(3) = 1: ArrayA(4) = 1: ArrayA(5) = 2
ArrayA(6) = 2: ArrayA(7) = 3: ArrayA(8) = 3: ArrayA(9) = 4

ArrayB(1) = 1: ArrayB(2) = 1: ArrayB(3) = 1: ArrayB(4) = 1: ArrayB(5) = 1
ArrayB(6) = 2: ArrayB(7) = 2: ArrayB(8) = 3: ArrayB(9) = 4

ArrayC(1) = 1: ArrayC(2) = 1: ArrayC(3) = 1: ArrayC(4) = 1: ArrayC(5) = 2
ArrayC(6) = 2: ArrayC(7) = 3: ArrayC(8) = 3: ArrayC(9) = 4
End Sub

Public Sub DeleteEntry(DinArrayCombSauArrayABC As Byte, poz As Byte, kt As Byte)
'poz=ce element din array tb stearsa, dupa idx
'kt=cate elemente sunt in vector
'DinArrayCombSauArrayABC=0 pentru arraycomb
'DinArrayCombSauArrayABC=1 pentru arraya
'DinArrayCombSauArrayABC=2 pentru arrayb
'DinArrayCombSauArrayABC=3 pentru arrayc
Dim i As Byte
Select Case DinArrayCombSauArrayABC
Case 0
    For i = poz + 1 To kt
        ArrayCombinatiiPunePiesa(i - 1) = ArrayCombinatiiPunePiesa(i)
    Next
Case 1
    For i = poz + 1 To kt
        ArrayA(i - 1) = ArrayA(i)
    Next
Case 2
    For i = poz + 1 To kt
        ArrayB(i - 1) = ArrayB(i)
    Next
Case 3
For i = poz + 1 To kt
        ArrayC(i - 1) = ArrayC(i)
    Next
End Select
End Sub

Private Function ResetPiesa(Ce As Piesa) As Piesa
'standard o piesa este 2222 si 0, adk nul tot
    Ce.ACuloare = 2
    Ce.BMarime = 2
    Ce.CForma = 2
    Ce.DSoliditate = 2
    Ce.ESett = 0
    ResetPiesa = Ce
End Function

Public Sub ResetTabla()
Dim i As Byte, j As Byte
For i = 1 To 4
    For j = 1 To 4
        Tabla(i, j) = ResetPiesa(Tabla(i, j))
    Next
Next
End Sub
Public Sub ResetArrayComb()
Dim i As Byte
For i = 1 To 15
    With ArrayCombinatiiPunePiesa(i)
    .LiniiDe2 = 0
    .LiniiDe3 = 0
    .NoComb = False
    .Block = 0
    .VictOnCuloare = 2
    .VictOnForma = 2
    .VictOnMarime = 2
    .VictOnSoliditate = 2
    .PozX = 0
    .PozY = 0
    .ArePiesa = True
    .CePiesaTreb.ACuloare = 2
    .CePiesaTreb.BMarime = 2
    .CePiesaTreb.CForma = 2
    .CePiesaTreb.DSoliditate = 2
    End With
Next
End Sub
Public Function CalculeazaComb(X As Byte, Y As Byte, Optional PentruFinal As Boolean = False) As Combinatie
Dim RezComb As Piesa 'rezultat la combinatie
Dim i As Byte, j As Byte, l As Byte, k As Byte 'counters
Dim kp As Byte 'cate piese pe linie/patrat
Dim l3 As Byte, l2 As Byte, V As Byte, NC As Boolean, Blk As Byte
    'linii de 3, de 2, victory, nocombination,block
Dim VoC As Byte, VoM As Byte, VoF As Byte, VoS  As Byte
    'victory on culoare, on marime, on forma, on soliditate
Dim CePiesaTreb As Piesa, Egz As Boolean

l3 = 0
l2 = 0
V = 0
NC = 0
Blk = 0
VoC = 2: VoM = 2: VoF = 2: VoS = 2

CePiesaTreb.ACuloare = 2
CePiesaTreb.BMarime = 2
CePiesaTreb.CForma = 2
CePiesaTreb.DSoliditate = 2
CePiesaTreb.ESett = 0
Egz = True 'exista
If PentruFinal Then 'cand afiseaza you lose sau you have won
    For i = 1 To 19
        ArrayVicts(i).ACuloare = 2
        ArrayVicts(i).BMarime = 2
        ArrayVicts(i).CForma = 2
        ArrayVicts(i).DSoliditate = 2
    Next
End If

'pentru linii orizontale

For i = 1 To 4
    kp = 0
    RezComb = ResetPiesa(RezComb)
    For j = 1 To 4
        If Tabla(i, j).ESett = 1 Then
            If RezComb.ESett = 0 Then 'prima piesa pe care o gaseste pe linie
                RezComb = Tabla(i, j)
            Else
                RezComb = CompleteazaRezComb(RezComb, i, j) 'o combina cu urmatoarele
            End If
        kp = kp + 1 'cate piese pe linie
        End If
    Next
    If RezComb.ESett = 1 Then 'dk pe linie a fost cel putin o piesa adk esett a fost preluat de la aceea
                'dk cel putin una din caract a fost gasita pe linie
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            Select Case kp
                Case 2
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then l2 = l2 + 1
                    If SingleOption(2) And RezComb.BMarime <> 2 Then l2 = l2 + 1
                    If SingleOption(3) And RezComb.CForma <> 2 Then l2 = l2 + 1
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then l2 = l2 + 1
                Case 3
                    'l3 = l3 + 1 ' are piesa
                    'dk sunt 3 piese tb ca cea pe care o va da sa nu contina o caract cu care sa invinga
                    '1-rc.(abcd)= dk este 0 devine 1, dk este 1 devine 0, adk inversul
                    If RezComb.ACuloare <> 2 And SingleOption(1) Then
                        'if-ul: dk a mai fost o data setarea, dar inversa inseamna ca piesa nu va exista
                        'cu siguranta (nu poate fi si alba si neagra in acelasi timp)
                        If CePiesaTreb.ACuloare = RezComb.ACuloare Then Egz = False
                        CePiesaTreb.ACuloare = 1 - RezComb.ACuloare
                        l3 = l3 + 1
                    End If
                    If RezComb.BMarime <> 2 And SingleOption(2) Then
                        If CePiesaTreb.BMarime = RezComb.BMarime Then Egz = False
                        CePiesaTreb.BMarime = 1 - RezComb.BMarime
                        l3 = l3 + 1
                    End If
                    If RezComb.CForma <> 2 And SingleOption(3) Then
                        If CePiesaTreb.CForma = RezComb.CForma Then Egz = False
                        CePiesaTreb.CForma = 1 - RezComb.CForma
                        l3 = l3 + 1
                    End If
                    If RezComb.DSoliditate <> 2 And SingleOption(4) Then
                        If CePiesaTreb.DSoliditate = RezComb.DSoliditate Then Egz = False
                        CePiesaTreb.DSoliditate = 1 - RezComb.DSoliditate
                        l3 = l3 + 1
                    End If
                    'egz=exista piesa care trebuie
                    'verifexistpiesa cauta prin piesele libere (vezi functia)
                    If Egz Then Egz = VerifExistaPiesa(CePiesaTreb)
                Case 4
                    'pt beginner: dk printre caract alese de utiliz nu se gasesc cele pe care ar putea fi
                    'victoria, nu va fi victorie (Vo(ABCD))
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then VoC = Tabla(X, Y).ACuloare: V = V + 1: If PentruFinal Then ArrayVicts(i).ACuloare = VoC
                    If SingleOption(2) And RezComb.BMarime <> 2 Then VoM = Tabla(X, Y).BMarime: V = V + 1: If PentruFinal Then ArrayVicts(i).BMarime = VoM
                    If SingleOption(3) And RezComb.CForma <> 2 Then VoF = Tabla(X, Y).CForma: V = V + 1: If PentruFinal Then ArrayVicts(i).CForma = VoF
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then VoS = Tabla(X, Y).DSoliditate: V = V + 1: If PentruFinal Then ArrayVicts(i).DSoliditate = VoS
            End Select
        End If
    End If
Next
'pentru linii verticale

For i = 1 To 4
    kp = 0
    RezComb = ResetPiesa(RezComb)
    For j = 1 To 4
        If Tabla(j, i).ESett = 1 Then
            If RezComb.ESett = 0 Then
                RezComb = Tabla(j, i)
            Else
                RezComb = CompleteazaRezComb(RezComb, j, i)
            End If
        kp = kp + 1
        End If
    Next
    If RezComb.ESett = 1 Then
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            Select Case kp
                Case 2
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then l2 = l2 + 1
                    If SingleOption(2) And RezComb.BMarime <> 2 Then l2 = l2 + 1
                    If SingleOption(3) And RezComb.CForma <> 2 Then l2 = l2 + 1
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then l2 = l2 + 1
                Case 3
                    
                    ' de aici verif are piesa
                    If RezComb.ACuloare <> 2 And SingleOption(1) Then
                        If CePiesaTreb.ACuloare = RezComb.ACuloare Then Egz = False
                        CePiesaTreb.ACuloare = 1 - RezComb.ACuloare
                        l3 = l3 + 1
                    End If
                    If RezComb.BMarime <> 2 And SingleOption(2) Then
                        If CePiesaTreb.BMarime = RezComb.BMarime Then Egz = False
                        CePiesaTreb.BMarime = 1 - RezComb.BMarime
                        l3 = l3 + 1
                    End If
                    If RezComb.CForma <> 2 And SingleOption(3) Then
                        If CePiesaTreb.CForma = RezComb.CForma Then Egz = False
                        CePiesaTreb.CForma = 1 - RezComb.CForma
                        l3 = l3 + 1
                    End If
                    If RezComb.DSoliditate <> 2 And SingleOption(4) Then
                        If CePiesaTreb.DSoliditate = RezComb.DSoliditate Then Egz = False
                        CePiesaTreb.DSoliditate = 1 - RezComb.DSoliditate
                        l3 = l3 + 1
                    End If
                    If Egz Then Egz = VerifExistaPiesa(CePiesaTreb)
                Case 4
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then VoC = Tabla(X, Y).ACuloare: V = V + 1: If PentruFinal Then ArrayVicts(4 + i).ACuloare = VoC
                    If SingleOption(2) And RezComb.BMarime <> 2 Then VoM = Tabla(X, Y).BMarime: V = V + 1: If PentruFinal Then ArrayVicts(4 + i).BMarime = VoM
                    If SingleOption(3) And RezComb.CForma <> 2 Then VoF = Tabla(X, Y).CForma: V = V + 1: If PentruFinal Then ArrayVicts(4 + i).CForma = VoF
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then VoS = Tabla(X, Y).DSoliditate: V = V + 1: If PentruFinal Then ArrayVicts(4 + i).DSoliditate = VoS
            End Select
        End If
    End If
Next
                    
'pentru diag princ
kp = 0
RezComb = ResetPiesa(RezComb)
For i = 1 To 4
        If Tabla(i, i).ESett = 1 Then
            If RezComb.ESett = 0 Then
                RezComb = Tabla(i, i)
            Else
                RezComb = CompleteazaRezComb(RezComb, i, i)
            End If
        kp = kp + 1
        End If
Next
    If RezComb.ESett = 1 Then
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            Select Case kp
                Case 2
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then l2 = l2 + 1
                    If SingleOption(2) And RezComb.BMarime <> 2 Then l2 = l2 + 1
                    If SingleOption(3) And RezComb.CForma <> 2 Then l2 = l2 + 1
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then l2 = l2 + 1
                Case 3
                     ' are piesa
                    If RezComb.ACuloare <> 2 And SingleOption(1) Then
                        If CePiesaTreb.ACuloare = RezComb.ACuloare Then Egz = False
                        CePiesaTreb.ACuloare = 1 - RezComb.ACuloare
                        l3 = l3 + 1
                    End If
                    If RezComb.BMarime <> 2 And SingleOption(2) Then
                        If CePiesaTreb.BMarime = RezComb.BMarime Then Egz = False
                        CePiesaTreb.BMarime = 1 - RezComb.BMarime
                        l3 = l3 + 1
                    End If
                    If RezComb.CForma <> 2 And SingleOption(3) Then
                        If CePiesaTreb.CForma = RezComb.CForma Then Egz = False
                        CePiesaTreb.CForma = 1 - RezComb.CForma
                        l3 = l3 + 1
                    End If
                    If RezComb.DSoliditate <> 2 And SingleOption(4) Then
                        If CePiesaTreb.DSoliditate = RezComb.DSoliditate Then Egz = False
                        CePiesaTreb.DSoliditate = 1 - RezComb.DSoliditate
                        l3 = l3 + 1
                    End If
                    If Egz Then Egz = VerifExistaPiesa(CePiesaTreb)
                Case 4
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then VoC = Tabla(X, Y).ACuloare: V = V + 1: If PentruFinal Then ArrayVicts(9).ACuloare = VoC
                    If SingleOption(2) And RezComb.BMarime <> 2 Then VoM = Tabla(X, Y).BMarime: V = V + 1: If PentruFinal Then ArrayVicts(9).BMarime = VoM
                    If SingleOption(3) And RezComb.CForma <> 2 Then VoF = Tabla(X, Y).CForma: V = V + 1: If PentruFinal Then ArrayVicts(9).CForma = VoF
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then VoS = Tabla(X, Y).DSoliditate: V = V + 1: If PentruFinal Then ArrayVicts(9).DSoliditate = VoS
            End Select
        End If
    End If
'Next

'pentru diag secundara
kp = 0
RezComb = ResetPiesa(RezComb)
For i = 1 To 4
    If Tabla(i, 5 - i).ESett = 1 Then
        If RezComb.ESett = 0 Then
            RezComb = Tabla(i, 5 - i)
        Else
            RezComb = CompleteazaRezComb(RezComb, i, 5 - i)
        End If
    kp = kp + 1
    End If
Next
    If RezComb.ESett = 1 Then
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            Select Case kp
                Case 2
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then l2 = l2 + 1
                    If SingleOption(2) And RezComb.BMarime <> 2 Then l2 = l2 + 1
                    If SingleOption(3) And RezComb.CForma <> 2 Then l2 = l2 + 1
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then l2 = l2 + 1
                Case 3
                     ' are piesa
                    If RezComb.ACuloare <> 2 And SingleOption(1) Then
                        If CePiesaTreb.ACuloare = RezComb.ACuloare Then Egz = False
                        CePiesaTreb.ACuloare = 1 - RezComb.ACuloare
                        l3 = l3 + 1
                    End If
                    If RezComb.BMarime <> 2 And SingleOption(2) Then
                        If CePiesaTreb.BMarime = RezComb.BMarime Then Egz = False
                        CePiesaTreb.BMarime = 1 - RezComb.BMarime
                        l3 = l3 + 1
                    End If
                    If RezComb.CForma <> 2 And SingleOption(3) Then
                        If CePiesaTreb.CForma = RezComb.CForma Then Egz = False
                        CePiesaTreb.CForma = 1 - RezComb.CForma
                        l3 = l3 + 1
                    End If
                    If RezComb.DSoliditate <> 2 And SingleOption(4) Then
                        If CePiesaTreb.DSoliditate = RezComb.DSoliditate Then Egz = False
                        CePiesaTreb.DSoliditate = 1 - RezComb.DSoliditate
                        l3 = l3 + 1
                    End If
                    If Egz Then Egz = VerifExistaPiesa(CePiesaTreb)
                Case 4
                    If SingleOption(1) And RezComb.ACuloare <> 2 Then VoC = Tabla(X, Y).ACuloare: V = V + 1: If PentruFinal Then ArrayVicts(10).ACuloare = VoC
                    If SingleOption(2) And RezComb.BMarime <> 2 Then VoM = Tabla(X, Y).BMarime: V = V + 1: If PentruFinal Then ArrayVicts(10).BMarime = VoM
                    If SingleOption(3) And RezComb.CForma <> 2 Then VoF = Tabla(X, Y).CForma: V = V + 1: If PentruFinal Then ArrayVicts(10).CForma = VoF
                    If SingleOption(4) And RezComb.DSoliditate <> 2 Then VoS = Tabla(X, Y).DSoliditate: V = V + 1: If PentruFinal Then ArrayVicts(10).DSoliditate = VoS
            End Select
        End If
    End If


'pentru patrate
If StilDeJoc = 3 Then 'hard=advanced
For i = 1 To 3
    For j = 1 To 3
    kp = 0
    RezComb = ResetPiesa(RezComb)
        For k = i To i + 1
            For l = j To j + 1
                If Tabla(k, l).ESett = 1 Then
                    If RezComb.ESett = 0 Then
                        RezComb = Tabla(k, l)
                    Else
                        RezComb = CompleteazaRezComb(RezComb, k, l)
                    End If
                kp = kp + 1
                End If
            Next
        Next
    If RezComb.ESett = 1 Then
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            Select Case kp
                'aici nu mai trebuie cu SingleOption(i) pentru ca stildejoc=3 deci singleoption(i)=true
                Case 2
                    If RezComb.ACuloare <> 2 Then l2 = l2 + 1
                    If RezComb.BMarime <> 2 Then l2 = l2 + 1
                    If RezComb.CForma <> 2 Then l2 = l2 + 1
                    If RezComb.DSoliditate <> 2 Then l2 = l2 + 1
                Case 3
                     ' are piesa
                    If RezComb.ACuloare <> 2 Then
                        If CePiesaTreb.ACuloare = RezComb.ACuloare Then Egz = False
                        CePiesaTreb.ACuloare = 1 - RezComb.ACuloare
                        l3 = l3 + 1
                    End If
                    If RezComb.BMarime <> 2 Then
                        If CePiesaTreb.BMarime = RezComb.BMarime Then Egz = False
                        CePiesaTreb.BMarime = 1 - RezComb.BMarime
                        l3 = l3 + 1
                    End If
                    If RezComb.CForma <> 2 Then
                        If CePiesaTreb.CForma = RezComb.CForma Then Egz = False
                        CePiesaTreb.CForma = 1 - RezComb.CForma
                        l3 = l3 + 1
                    End If
                    If RezComb.DSoliditate <> 2 Then
                        If CePiesaTreb.DSoliditate = RezComb.DSoliditate Then Egz = False
                        CePiesaTreb.DSoliditate = 1 - RezComb.DSoliditate
                        l3 = l3 + 1
                    End If
                    If Egz Then Egz = VerifExistaPiesa(CePiesaTreb)
                Case 4
                    If RezComb.ACuloare <> 2 Then VoC = Tabla(X, Y).ACuloare: V = V + 1: If PentruFinal Then ArrayVicts(10 + (i - 1) * 3 + j).ACuloare = VoC
                    If RezComb.BMarime <> 2 Then VoM = Tabla(X, Y).BMarime: V = V + 1: If PentruFinal Then ArrayVicts(10 + (i - 1) * 3 + j).BMarime = VoM
                    If RezComb.CForma <> 2 Then VoF = Tabla(X, Y).CForma: V = V + 1: If PentruFinal Then ArrayVicts(10 + (i - 1) * 3 + j).CForma = VoF
                    If RezComb.DSoliditate <> 2 Then VoS = Tabla(X, Y).DSoliditate: V = V + 1: If PentruFinal Then ArrayVicts(10 + (i - 1) * 3 + j).DSoliditate = VoS
            End Select
        End If
    End If
    Next
Next
End If

''''''''''''''''''''''''''''''''''''''''''
'block
Blk = 0
'linie
RezComb = ResetPiesa(RezComb)
For i = 1 To 4
    If Y <> i Then
        If Tabla(X, i).ESett = 1 Then
            If RezComb.ESett = 0 Then
                RezComb = Tabla(X, i)
            Else
                RezComb = CompleteazaRezComb(RezComb, X, i)
            End If
        End If
    End If
Next
'dk macar o caract este posibila pe linie...
If (RezComb.ACuloare <> 2 And SingleOption(1)) Or (RezComb.BMarime <> 2 And SingleOption(2)) _
Or (RezComb.CForma <> 2 And SingleOption(3)) Or (RezComb.DSoliditate <> 2 And SingleOption(4)) Then
    '... si dupa combinatia cu piesa pusa!!! nu mai este atunci inseamna ca a blocat linia!
    RezComb = CompleteazaRezComb(RezComb, X, Y)
    If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
        Blk = Blk + 1
    End If
End If

'coloana
RezComb = ResetPiesa(RezComb)
For i = 1 To 4
    If i <> X Then
        If Tabla(i, Y).ESett = 1 Then
            If RezComb.ESett = 0 Then
                RezComb = Tabla(i, Y)
            Else
                RezComb = CompleteazaRezComb(RezComb, i, Y)
            End If
        End If
    End If
Next
If (RezComb.ACuloare <> 2 And SingleOption(1)) Or (RezComb.BMarime <> 2 And SingleOption(2)) _
Or (RezComb.CForma <> 2 And SingleOption(3)) Or (RezComb.DSoliditate <> 2 And SingleOption(4)) Then
    RezComb = CompleteazaRezComb(RezComb, X, Y)
    If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
        Blk = Blk + 1
    End If
End If

'diag princ
RezComb = ResetPiesa(RezComb)
If X = Y Then
    For i = 1 To 4
    If i <> X Then
        If Tabla(i, i).ESett = 1 Then
            If RezComb.ESett = 0 Then
                RezComb = Tabla(i, i)
            Else
                RezComb = CompleteazaRezComb(RezComb, i, i)
            End If
        End If
    End If
    Next
    If (RezComb.ACuloare <> 2 And SingleOption(1)) Or (RezComb.BMarime <> 2 And SingleOption(2)) _
    Or (RezComb.CForma <> 2 And SingleOption(3)) Or (RezComb.DSoliditate <> 2 And SingleOption(4)) Then
        RezComb = CompleteazaRezComb(RezComb, X, Y)
        If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
            Blk = Blk + 1
        End If
    End If
End If

'diag sec
RezComb = ResetPiesa(RezComb)
If X + Y = 5 Then
    For i = 1 To 4
    If i <> X And 5 - i <> Y Then
        If Tabla(i, 5 - i).ESett = 1 Then
            If RezComb.ESett = 0 Then
                RezComb = Tabla(i, 5 - i)
            Else
                RezComb = CompleteazaRezComb(RezComb, i, 5 - i)
            End If
        End If
    End If
    Next
    If (RezComb.ACuloare <> 2 And SingleOption(1)) Or (RezComb.BMarime <> 2 And SingleOption(2)) _
    Or (RezComb.CForma <> 2 And SingleOption(3)) Or (RezComb.DSoliditate <> 2 And SingleOption(4)) Then
        RezComb = CompleteazaRezComb(RezComb, X, Y)
        If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
            Blk = Blk + 1
        End If
    End If
End If

If StilDeJoc = 3 Then
    
    'patrat dr jos
    If X <= 3 And Y <= 3 Then
    RezComb = ResetPiesa(RezComb)
        For i = X To X + 1
            For j = Y To Y + 1
                If i <> X Or Y <> j Then
                    If Tabla(i, j).ESett = 1 Then
                        If RezComb.ESett = 0 Then
                            RezComb = Tabla(i, j)
                        Else
                            RezComb = CompleteazaRezComb(RezComb, i, j)
                        End If
                    End If
                End If
            Next
        Next
        'aici nu mai trebuie cu SingleOption(i) pentru ca stildejoc=3 deci singleoption(i)=true
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            RezComb = CompleteazaRezComb(RezComb, X, Y)
            If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
                Blk = Blk + 1
            End If
        End If
    End If
    'patrat dr sus
    If X >= 2 And Y <= 3 Then
    RezComb = ResetPiesa(RezComb)
        For i = X - 1 To X
            For j = Y To Y + 1
                If i <> X Or Y <> j Then
                    If Tabla(i, j).ESett = 1 Then
                        If RezComb.ESett = 0 Then
                            RezComb = Tabla(i, j)
                        Else
                            RezComb = CompleteazaRezComb(RezComb, i, j)
                        End If
                    End If
                End If
            Next
        Next
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            RezComb = CompleteazaRezComb(RezComb, X, Y)
            If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
                Blk = Blk + 1
            End If
        End If
    End If
    'patrat st jos
    If X <= 3 And Y >= 2 Then
    RezComb = ResetPiesa(RezComb)
        For i = X To X + 1
            For j = Y - 1 To Y
                If i <> X Or Y <> j Then
                    If Tabla(i, j).ESett = 1 Then
                        If RezComb.ESett = 0 Then
                            RezComb = Tabla(i, j)
                        Else
                            RezComb = CompleteazaRezComb(RezComb, i, j)
                        End If
                    End If
                End If
            Next
        Next
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            RezComb = CompleteazaRezComb(RezComb, X, Y)
            If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
                Blk = Blk + 1
            End If
        End If
    End If
    'patrat st sus
    If X >= 2 And Y >= 2 Then
    RezComb = ResetPiesa(RezComb)
        For i = X - 1 To X
            For j = Y - 1 To Y
                If i <> X Or Y <> j Then
                    If Tabla(i, j).ESett = 1 Then
                        If RezComb.ESett = 0 Then
                            RezComb = Tabla(i, j)
                        Else
                            RezComb = CompleteazaRezComb(RezComb, i, j)
                        End If
                    End If
                End If
            Next
        Next
        If RezComb.ACuloare <> 2 Or RezComb.BMarime <> 2 Or RezComb.CForma <> 2 Or RezComb.DSoliditate <> 2 Then
            RezComb = CompleteazaRezComb(RezComb, X, Y)
            If RezComb.ACuloare = 2 And RezComb.BMarime = 2 And RezComb.CForma = 2 And RezComb.DSoliditate = 2 Then
                Blk = Blk + 1
            End If
        End If
    End If
End If
If l3 = 0 And l2 = 0 And V = 0 And Blk = 0 Then
    NC = True
End If
'transmite pt vector
CalculeazaComb.ArePiesa = Egz
CalculeazaComb.Block = Blk
CalculeazaComb.LiniiDe2 = l2
CalculeazaComb.LiniiDe3 = l3
CalculeazaComb.NoComb = NC
CalculeazaComb.PozX = X
CalculeazaComb.PozY = Y
CalculeazaComb.VictOnCuloare = VoC
CalculeazaComb.VictOnForma = VoF
CalculeazaComb.VictOnMarime = VoM
CalculeazaComb.VictOnSoliditate = VoS
CalculeazaComb.Victory = V
CalculeazaComb.CePiesaTreb = CePiesaTreb
End Function

Private Function CompleteazaRezComb(Ce As Piesa, X As Byte, Y As Byte) As Piesa
'dk intre rezcomb si piesa cu care face combinatia nu exitsa aceeasi caract atunci rc devine nul:
'caracteristica aceea nu exista pe linie
    If Ce.ACuloare <> Tabla(X, Y).ACuloare Then Ce.ACuloare = 2
    
    If Ce.BMarime <> Tabla(X, Y).BMarime Then Ce.BMarime = 2
    
    If Ce.CForma <> Tabla(X, Y).CForma Then Ce.CForma = 2
    
    If Ce.DSoliditate <> Tabla(X, Y).DSoliditate Then Ce.DSoliditate = 2
    
    CompleteazaRezComb = Ce
End Function
Private Function VerifExistaPiesa(Ce As Piesa) As Boolean 'rc=rezcomb
Dim i As Byte
Dim Temp As Piesa
For i = 1 To 16
    If Piese(i).ESett = 0 Then
        'temp=o copie a fiecarei piese libere din toate piesele
        Temp = Piese(i)
        'modifica copia astfel incat valorile nule ale lui Rezcomb sa existe si in ea(copie)
        If Ce.ACuloare = 2 Then Temp.ACuloare = 2
        If Ce.BMarime = 2 Then Temp.BMarime = 2
        If Ce.CForma = 2 Then Temp.CForma = 2
        If Ce.DSoliditate = 2 Then Temp.DSoliditate = 2
        'apoi face comparatia
        'dk sunt identice(copia=temp si rezcomb) atunci piesa cautata exista
        'iese din for pentru ca este esential dk este cel putin una=bool
        If Temp.ACuloare = Ce.ACuloare And Temp.BMarime = Ce.BMarime And _
        Temp.CForma = Ce.CForma And Temp.DSoliditate = Ce.DSoliditate Then
            VerifExistaPiesa = True
            Exit For
        End If
    End If
Next
End Function

Public Function VerifCombLaFel(Prima As Combinatie, ADoua As Combinatie)

If Prima.ArePiesa = ADoua.ArePiesa And _
    Prima.Block = ADoua.Block And _
    Prima.CePiesaTreb.ACuloare = ADoua.CePiesaTreb.ACuloare And _
    Prima.CePiesaTreb.BMarime = ADoua.CePiesaTreb.BMarime And _
    Prima.CePiesaTreb.CForma = ADoua.CePiesaTreb.CForma And _
    Prima.CePiesaTreb.DSoliditate = ADoua.CePiesaTreb.DSoliditate And _
    Prima.CePiesaTreb.ESett = ADoua.CePiesaTreb.ESett And _
    Prima.LiniiDe2 = ADoua.LiniiDe2 And _
    Prima.LiniiDe3 = ADoua.LiniiDe3 And _
    Prima.NoComb = ADoua.NoComb And _
    Prima.VictOnCuloare = ADoua.VictOnCuloare And _
    Prima.VictOnForma = ADoua.VictOnForma And _
    Prima.VictOnMarime = ADoua.VictOnMarime And _
    Prima.VictOnSoliditate = ADoua.VictOnSoliditate And _
    Prima.Victory = ADoua.Victory Then
    
    VerifCombLaFel = True
Else
    VerifCombLaFel = False
End If
End Function

Public Sub CreeazaArrayPiese() 'se apeleaza o singura data pe joc
Dim i As Byte, j As Byte, k As Byte, l As Byte
Dim PiesaC As Piesa 'piesa pe care o introd in vector
Dim kp As Byte 'nr de piese=idx de vector
'fiecare piesa face parte din sirul 0000,0001,0010,0011,etc generat de 4 foruri
kp = 0
For i = 0 To 1
    For j = 0 To 1
        For k = 0 To 1
            For l = 0 To 1
                With PiesaC
                    .ACuloare = i
                    .BMarime = j
                    .CForma = k
                    .DSoliditate = l
                    .ESett = 0
                End With
                kp = kp + 1
                Piese(kp) = PiesaC
            Next
        Next
    Next
Next

End Sub
