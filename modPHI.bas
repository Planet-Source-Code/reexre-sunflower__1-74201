Attribute VB_Name = "modPHI"
'************************************************
' PHI number, Golden Ration
' By reexre
'*************************************************


Option Explicit

Private Type tP
    x              As Double
    y              As Double
    tX             As Double
    tY             As Double
    R              As Double
    tR              As Double
    U              As Boolean
End Type

Public P()         As tP

Private Const PI   As Double = 3.14159265358979
Private Const PI2  As Double = 6.28318530717959

Private Const MyC  As Long = 3338495    'rgb(255,240,50)

Public PHI         As Double

Private x          As Double
Private y          As Double
Public cX          As Double
Public cY          As Double

Public ExitLoop    As Boolean

Private CNT        As Long
Private FR         As Long


Public Sub FULLDraw(ByRef PIC As PictureBox)

    Dim I          As Long
    Dim Scala      As Double
    Dim A          As Double
    Dim hDC        As Long
    Dim N          As Long

    hDC = PIC.hDC


    N = 1500


    PHI = (1 + Sqr(5)) / 2 - 1
    Scala = ((cY - 20) / Sqr(N))


    Dim SA         As Double

    For SA = 0 To PI2 Step PI2 * 0.05

        BitBlt hDC, 0, 0, cX * 2 \ 1, cY * 2 \ 1, hDC, 0, 0, vbBlack
        A = SA
        For I = 1 To N
            x = cX + Sqr(I) * Cos(A) * Scala
            y = cY + Sqr(I) * Sin(A) * Scala
            A = A + PHI * PI2
            MyCircle hDC, x \ 1, y \ 1, 2, 2, vbRed
        Next
        PIC.Refresh

    Next
End Sub


Public Sub Animation(ByRef PIC As PictureBox, NP As Long, Speed As Double)
    Dim hDC        As Long

    Dim N          As Long
    Dim I          As Long
    Dim J          As Long
    Dim Scala      As Double
    Dim K          As Long
    Dim Kto        As Long
    Dim A          As Double
    Dim D          As Double


    Dim K1 As Double
    Dim K2 As Double
    
    K1 = 0.95 '0.97
    K2 = 1 - K1


    If NP < 3 Then NP = 3
    N = NP
    ReDim P(N)
    For I = 1 To N
        P(I).x = cX
        P(I).y = cY
    Next

    hDC = PIC.hDC
    PHI = (1 + Sqr(5)) / 2 - 1
    Scala = ((cY - 20) / Sqr(N))

    frmMain.MousePointer = 11

    CNT = 0
    FR = 0
    For I = 1 To N
        'Scala = ((cY - 20) / Sqr(I))
        Scala = ((cY - 20) / (I) ^ PHI)
        '------------------------------------
        For J = 1 To I
            'D = Sqr(I - J + 1)
            D = (I - J + 1) ^ PHI
            '----------------------------------------

            A = PHI * J * PI2
            P(J).tX = cX + D * Scala * Cos(A)
            P(J).tY = cY + D * Scala * Sin(A)
            P(J).tR = 0.95 * (1 - PHI) * Scala / I * (I * I - J * J) ^ PHI

        Next

        Kto = 0.5 * (1 - I / N) * (101 - Speed)

        If I Mod 10 = 0 Then
            frmMain.Caption = I & " / " & N & "   " & Int(100 * I / N) & "%  (" & Kto & ")"
            DoEvents
        End If


        For K = 1 To Kto
            For J = 1 To I
                P(J).x = P(J).x * K1 + P(J).tX * K2
                P(J).y = P(J).y * K1 + P(J).tY * K2
                P(J).R = P(J).R * K1 + P(J).tR * K2
            Next
            BitBlt hDC, 0, 0, cX * 2 \ 1, cY * 2 \ 1, hDC, 0, 0, vbBlack
            If frmMain.chRad Then
                For J = 1 To I
                    MyCircle hDC, P(J).x \ 1, P(J).y \ 1, P(J).R * 0.5 + 1, (P(J).R + 1) \ 1, MyC
                Next
            Else
                For J = 1 To I
                    MyCircle hDC, P(J).x \ 1, P(J).y \ 1, 2, 2, MyC
                Next
            End If


            PIC.Refresh

            SaveFrame



            If ExitLoop Then I = N: K = N + 9999
        Next



    Next
    '-----------------------------------
    If ExitLoop Then
        I = J - 1
    Else
        I = N
    End If


    For K = 1 To 200
        For J = 1 To I
            P(J).x = P(J).x * K1 + P(J).tX * K2
            P(J).y = P(J).y * K1 + P(J).tY * K2
             P(J).R = P(J).R * K1 + P(J).tR * K2
        Next
        BitBlt hDC, 0, 0, cX * 2 \ 1, cY * 2 \ 1, hDC, 0, 0, vbBlack
        If frmMain.chRad Then
            For J = 1 To I
                  MyCircle hDC, P(J).x \ 1, P(J).y \ 1, P(J).R * 0.5 + 1, (P(J).R + 1) \ 1, MyC
            Next
        Else
            For J = 1 To I
                MyCircle hDC, P(J).x \ 1, P(J).y \ 1, 2, 2, MyC
            Next
        End If

        PIC.Refresh
        DoEvents
        SaveFrame
    Next

    'Join I

    ExitLoop = False
    frmMain.Command1.Enabled = True
    frmMain.MousePointer = 0


End Sub


Public Sub Join(NN As Long)
    Dim I          As Long
    Dim Dmin       As Double
    Dim D          As Double
    Dim J          As Long

    Dim cp         As Long
    Dim nV         As Long
    Dim Cent       As tP

    Cent.tX = cX
    Cent.tY = cY

    cp = NN
    P(cp).U = True
AG:

    Dmin = 999999999999#
    For I = 1 To NN
        If I <> cp Then
            If P(I).U = False Then
                D = DistSQ(P(I), P(cp))    '+ DistSQ(P(I), Cent) * (0.025 * PHI)

                If D < Dmin Then
                    Dmin = D
                    J = I
                End If
            End If
        End If
    Next

    P(J).U = True
    FastLine frmMain.PIC.hDC, P(J).x \ 1, P(J).y \ 1, P(cp).x \ 1, P(cp).y \ 1, 1, vbCyan
    frmMain.PIC.Refresh

    cp = J
    nV = nV + 1
    If nV < NN Then GoTo AG





End Sub
Private Function DistSQ(P1 As tP, P2 As tP) As Double
    Dim Dx         As Double
    Dim Dy         As Double
    Dx = P2.tX - P1.tX
    Dy = P2.tY - P1.tY
    DistSQ = Dx * Dx + Dy * Dy
End Function


Public Sub SaveFrame()
    CNT = CNT + 1
    If frmMain.chJpg Then
        If CNT Mod 3 = 0 Then
            SaveJPG frmMain.PIC.Image, App.Path & "\Frames\" & Format(FR, "00000") & ".jpg", 94
            FR = FR + 1
            frmMain.Label1.Caption = FR & "   " & Int(FR / 30)
            DoEvents
        End If
    End If

End Sub
