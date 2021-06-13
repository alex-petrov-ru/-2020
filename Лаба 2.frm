VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   16245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Разгонка"
      Height          =   1095
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Загонка"
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   9015
      Left            =   3480
      ScaleHeight     =   8955
      ScaleWidth      =   12195
      TabIndex        =   0
      Top             =   120
      Width           =   12255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C(), Ck(), X(), y(), a(), b(), r(), d(), del(), lam(), Q, T, Rp, delRp, dC, dCk, dd As Single

Function dif(C, dd) '' Функция для рассчета коэффициента диффузии ''
k = 0.0000861
Nc = 4.831E+15 * ((1.08 * T) ^ 1.5)
Nv = 4.831E+15 * ((0.59 * T) ^ 1.5)
Eg = 1.21 - (0.00028) * T
ni = ((Nc * Nv / 2) ^ 0.5) * Exp(-Eg / (2 * k * T))
dif = (0.214 * Exp(-3.65 / (k * T)) + 15 * (C / ni) * Exp(-4.08 / (k * T))) * 2 * dd
End Function

Function f(X) '' Гауссово распределение, на входе X в мкм, в процессе расчета переводятся в см ''
f = Q / (2.5 * delRp) * Exp(-(X * 10 ^ -4 - Rp) ^ 2 / (2 * delRp ^ 2))
End Function
Function erfic(a)
erfic = 1 - (1 / (1 + 0.278393 * a + 0.230389 * a ^ 2 + 0.000972 * a ^ 3 + 0.078108 * a ^ 4) ^ 4)
End Function

Function osi() 'Рисует оси'
'Рисуем ось Х - горизонтальная'
Y1 = 2 / 3 * Picture1.ScaleHeight
X1 = 1 / 3 * Picture1.ScaleWidth
Y2 = 2 / 3 * Picture1.ScaleHeight
X2 = Picture1.ScaleWidth
Picture1.Line (X1, Y1)-(X2, Y2)
'Рисуем ось У - под наклоном'
Y3 = Picture1.ScaleHeight
X3 = 0
Picture1.Line (X1, Y1)-(X3, Y3)
'Рисуем ось С - вертикальная'
X4 = 1 / 3 * Picture1.ScaleWidth
Y4 = 0
Picture1.Line (X1, Y1)-(X4, Y4)
'Рисуем отсечки. Принимаем длину У = 4 мкм, размер окна - 2 мкм на расстоянии 1 мкм от начала и конца оси У'
'Отсечка по 3/4 У'
X11 = (X1 * 1 / 4 - 300)
Y11 = (Y1 + Y3 * 3 / 12)
X21 = X11 + 600
Y21 = Y11
Picture1.Line (X11, Y11)-(X21, Y21)
'Отсечка по 1/4 У'
X12 = (X1 * 3 / 4 - 300)
Y12 = (Y1 + Y3 * 1 / 12)
X22 = X12 + 600
Y22 = Y12
Picture1.Line (X12, Y12)-(X22, Y22)
End Function

Private Sub Command1_Click()
'' Исходные данные ''
Q = 500000000000000#
E = 180000
T = 1273
'' Материал - Sb, подложка Si<111> p-тип, Na = 1е16 ''
Rp = (0.000668 * E ^ 0.921 + 0.005072) / 10000000# '' 46.23 нм, переводятся в см, деление на 10е-7''
delRp = (0.000241 * E ^ 0.884 + 0.000923) / 10000000# '' 10.16 нм, переводятся в см, деление на 10е-7''
osi
'Начинаем заполнять массивы'
n = 1000
m = 50
ReDim C(n + 1, m + 1), Ck(n + 1, m + 1), X(n + 1), y(m + 1), a(n + 1), b(n + 1), r(n + 1), d(n + 1), del(n + 1), lam(n + 1)
X00 = 1 / 3 * Picture1.ScaleWidth
Y00 = 2 / 3 * Picture1.ScaleHeight
Xlast = 0
Ylast = Picture1.ScaleHeight
dxx = Picture1.ScaleWidth / (n - 1)
dyy = (Ylast - Y00) / (m - 1)
'Задать х(i) и у(i)'
Xmax = 1
dx = Xmax / n
Ymax = 4
dy = Ymax / m
X(1) = 0
y(1) = 0
For j = 2 To m
    For i = 2 To n
        X(i) = X(i - 1) + dx
    Next i
    y(j) = y(j - 1) + dy
Next j

'Начинаем заполнять функцию'
dRPL = delRp / 2
For j = 1 To m
    For i = 1 To n
        If (y(j) <= 1) Or (y(j) >= 3) Then
            C(i, j) = f(X(i)) * 1 / 2 * (erfic((y(j) - 2) / (1.414 * dRPL)) - erfic((y(j) + 2) / (1.414 * dRPL)))
        Else
            C(i, j) = f(X(i))
        End If
    Next i
Next j
For j = 1 To m - 1
    For i = 1 To n
            Picture1.PSet (X(i) / Xmax * 2 / 3 * Picture1.ScaleWidth + 1 / 3 * Picture1.ScaleWidth - 1 / 3 * dxx * j * n / m,
            Picture1.ScaleHeight - C(i, j) / 3E+20 * Picture1.ScaleHeight - 1 / 3 * Picture1.ScaleHeight + dyy * j)
    Next i
Next j

End Sub

Private Sub Command2_Click()
'' Исходные данные ''
Q = 500000000000000#
E = 180000
T = 1273
'' Материал - Sb, подложка Si<111> p-тип, Na = 1е16 ''
Rp = (0.000668 * E ^ 0.921 + 0.005072) / 10000000# '' 46.23 нм, переводятся в см, деление на 10е-7''
delRp = (0.000241 * E ^ 0.884 + 0.000923) / 10000000# '' 10.16 нм, переводятся в см, деление на 10е-7''
Picture1.Cls
osi
ke = 1.9E-61
'Начинаем заполнять массивы'
n = 1000
m = 50
ReDim C(n + 1, m + 1), Ck(n + 1, m + 1), X(n + 1), y(m + 1), a(n + 1), b(n + 1), r(n + 1), d(n + 1), del(n + 1), lam(n + 1)
X00 = 1 / 3 * Picture1.ScaleWidth
Y00 = 2 / 3 * Picture1.ScaleHeight
Xlast = 0
Ylast = Picture1.ScaleHeight
dxx = Picture1.ScaleWidth / (n - 1)
dyy = (Ylast - Y00) / (m - 1)
'Задать х(i) и у(i)'
Xmax = 1
dx = Xmax / n
Ymax = 4
dy = Ymax / m
X(1) = 0
y(1) = 0
For j = 2 To m
    For i = 2 To n
        X(i) = X(i - 1) + dx
    Next i
    y(j) = y(j - 1) + dy
Next j
'Начинаем заполнять функцию'
dRPL = delRp / 2
For j = 1 To m
    For i = 1 To n
        If (y(j) <= 1) Or (y(j) >= 3) Then
            C(i, j) = f(X(i)) * 1 / 2 * (erfic((y(j) - 2) / (1.414 * dRPL)) - erfic((y(j) + 2) / (1.414 * dRPL)))
        Else
            C(i, j) = f(X(i))
        End If
        Ck(i, j) = (C(i, j) / (3 * ke)) ^ 0.25
    Next i
Next j
dt = 90
tt = 900
'' Расчет разгонки, часть полушага, отвечающая за координату X''
b(1) = 0
a(1) = -1
d(1) = 1
r(1) = 0
del(1) = -d(1) / a(1)
lam(1) = r(1) / a(1)
b(n) = 0
d(n) = 0
a(n) = 1
r(n) = 0
For g = 0 To tt Step dt
    Picture1.Cls
    osi
    For j = 1 To m - 1
        For i = 2 To n
            If j > 3 / 4 * m And g = 0 Then
            C(i, j) = 0
            End If
            dC = C(i + 1, j) - C(i, j)
            dCk = Ck(i + 1, j) - Ck(i, j)
            If dC < 0.01 Or dCk < 0.01 Then '' Если dC или dCk близки к нулю, выдается ошибка в программе ''
                dd = 1
            Else
                dd = dCk / dC
            End If
            a(i) = -(2 + ((dx ^ 2) * 0.00000001 / (dif(C(i, j), dd) * dt)))
            b(i) = 1
            d(i) = 1
            r(i) = -((dx ^ 2) * 0.00000001 * C(i, j)) / (dif(C(i, j), dd) * dt)
        Next i
        del(1) = -d(1) / a(1)
        lam(1) = r(1) / a(1)
        For i = 2 To n - 1
            del(i) = -d(i) / (a(i) + b(i) * del(i - 1))
            lam(i) = (r(i) - b(i) * lam(i - 1)) / (a(i) + b(i) * del(i - 1))
        Next i
        C(n, j) = lam(n)
        For i = n - 1 To 1 Step -1
                C(i, j) = del(i) * C(i + 1, j) + Abs(lam(i))
        Next i
    Next j
    For i = 1 To n - 1 '' Расчет разгонки примеси, часть полушага, отвечающая за координату У''
        For j = 2 To m - 1
            If j > 3 / 4 * m And g = 0 Then
                C(i, j) = 0
            End If
            dC = C(i, j + 1) - C(i, j + 1)
            dCk = Ck(i, j + 1) - Ck(i, j + 1)
            If dC < 0.01 Or dCk < 0.01 Then '' Если dC или dCk близки к нулю, выдается ошибка в программе ''
                dd = 1
            Else
                dd = dCk / dC
            End If
            a(j) = -(2 + ((dy ^ 2) * 0.00000001 / (dif(C(i, j), dd) * dt)))
            b(j) = 1
            d(j) = 1
            r(j) = -((dy ^ 2) / (dx ^ 2) * 0.00000001 * (C(i + 1, j) - (2 - (dx ^ 2 / (dif(C(i, j), dd) * dt))) * C(i, j) + C(i - 1, j)))
        Next j
        del(1) = -d(1) / a(1)
        lam(1) = r(1) / a(1)
        For j = 2 To m - 1
            del(j) = -d(j) / (a(j) + b(j) * del(j - 1))
            lam(j) = (r(j) - b(j) * lam(j - 1)) / (a(j) + b(j) * del(j - 1))
        Next j
        C(i, m) = lam(m)
        For j = m - 1 To 1 Step -1
                C(i, j) = del(j) * C(i, j + 1) + Abs(lam(j))
        Next j
    Next i
    For j = 1 To m - 1
        For i = 1 To n
            Picture1.PSet (X(i) / Xmax * 2 / 3 * Picture1.ScaleWidth + 1 / 3 * Picture1.ScaleWidth - 1 / 3 * dxx * j * n / m,
            Picture1.ScaleHeight - C(i, j) / 3E+20 * Picture1.ScaleHeight - 1 / 3 * Picture1.ScaleHeight + dyy * j)
        Next i
    Next j
Next g
End Sub

