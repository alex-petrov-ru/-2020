VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17475
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   17475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   240
      TabIndex        =   4
      Text            =   "Глубина залегания p-n перехода, мкм "
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Разгонка"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   8535
      Left            =   3480
      ScaleHeight     =   8475
      ScaleWidth      =   13635
      TabIndex        =   1
      Top             =   360
      Width           =   13695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Загонка"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C(), Ck(), X(), a(), b(), d(), r(), del(), lam(), Q, T, Rp, delRp, dC, dCk, dd As Single

Function f(X) '' Гауссово распределение, на входе X в мкм, в процессе расчета переводятся в см ''
f = Q / (2.5 * delRp) * Exp(-(X * 10 ^ -4 - Rp) ^ 2 / (2 * delRp ^ 2))
End Function

Function dif(C, dd) '' Функция для рассчета коэффициента диффузии ''
k = 0.0000861
Nc = 4.831E+15 * ((1.08 * T) ^ 1.5)
Nv = 4.831E+15 * ((0.59 * T) ^ 1.5)
Eg = 1.21 - (0.00028) * T
ni = ((Nc * Nv / 2) ^ 0.5) * Exp(-Eg / (2 * k * T))
dif = (0.214 * Exp(-3.65 / (k * T)) + 15 * (C / ni) * Exp(-4.08 / (k * T))) * 2 * dd
End Function

Private Sub Command1_Click()
'' Исходные данные ''
Q = 500000000000000#
E = 180000
n = 3000
T = 1273
'' Материал - Sb, подложка Si<111> p-тип, Na = 1е16 ''
Rp = (0.000668 * E ^ 0.921 + 0.005072) / 10000000# '' 46.23 нм, переводятся в см, деление на 10е-7''
delRp = (0.000241 * E ^ 0.884 + 0.000923) / 10000000# '' 10.16 нм, переводятся в см, деление на 10е-7''
ReDim C(n + 1), X(n + 1), Ck(n + 1)
Xmax = 1
dx = Xmax / (n - 1)
X(1) = 0
C(1) = f(X(1))
For i = 2 To n - 1 '' Заполнение массива концентрации через функцию распределения Гаусса ''
    X(i) = X(i - 1) + dx
    C(i) = f(X(i))
Next i
For i = 1 To n '' Вывод распределения на экран ''
    Picture1.PSet (X(i) / Xmax * Picture1.ScaleWidth, Picture1.ScaleHeight - C(i) / 2E+20 * Picture1.ScaleHeight)
Next i
End Sub

Private Sub Command2_Click()
n = 3000
Q = 500000000000000#
E = 180000
tt = 900 '' Время ''
T = 1273
pp = 1E+16 '' Концентрация ОНЗ в подложке ''
ke = 1.9E-61
Rp = (0.000668 * E ^ 0.921 + 0.005072) / 10000000# '' 46.23 нм, переводятся в см, деление на 10е-7''
delRp = (0.000241 * E ^ 0.884 + 0.000923) / 10000000# '' 10.16 нм, переводятся в см, деление на 10е-7''
ReDim C(n + 1), Ck(n + 1), X(n + 1), a(n + 1), b(n + 1), d(n + 1), r(n + 1), del(n + 1), lam(n + 1)
Xmax = 1
dx = Xmax / (n - 1)
X(1) = 0
C(1) = f(X(1))
For i = 2 To n
    X(i) = X(i - 1) + dx
    C(i) = f(X(i))
    Ck(i) = (C(i) / (3 * ke)) ^ 0.25 '' Концентрация для учета эффекта кластеризации ''
Next i
dt = 1
b(1) = 0
a(1) = -1
d(1) = 1
r(1) = 0
For g = 1 To tt
    Picture1.Cls
    For i = 2 To n - 1
        dC = C(i + 1) - C(i)
        dCk = Ck(i + 1) - Ck(i)
        If dC < 0.01 Or dCk < 0.01 Then '' Если dC или dCk близки к нулю, выдается ошибка в программе ''
            dd = 1
        Else
            dd = dCk / dC
        End If
        a(i) = -(2 + ((dx ^ 2) * 0.00000001 / (dif(C(i), dd) * dt)))
        b(i) = 1
        d(i) = 1
        r(i) = -((dx ^ 2) * 0.00000001 * C(i)) / (dif(C(i), dd) * dt)
    Next i
    del(1) = -d(1) / a(1)
    lam(1) = r(1) / a(1)
    For i = 2 To n - 1
        del(i) = -d(i) / (a(i) + b(i) * del(i - 1))
        lam(i) = (r(i) - b(i) * lam(i - 1)) / (a(i) + b(i) * del(i - 1))
    Next i
    C(n) = lam(n)
    For i = n - 1 To 1 Step -1
        C(i) = del(i) * C(i + 1) + Abs(lam(i))
    Next i
    For i = 1 To n
        If g > tt * 0.95 Then '' Конечное распределение - синий цвет, промежуточные - пурпурный ''
            Picture1.PSet (X(i) / Xmax * Picture1.ScaleWidth, Picture1.ScaleHeight - C(i) / 2E+20 * Picture1.ScaleHeight), vbBlue
        Else
            Picture1.PSet (X(i) / Xmax * Picture1.ScaleWidth, Picture1.ScaleHeight - C(i) / 2E+20 * Picture1.ScaleHeight), vbMagenta
        End If
    Next i
    For i = 1 To n '' Определение глубины p-n перехода ''
        If C(i) > pp And C(i + 1) < pp Then
            Text1 = X(i)
        End If
    Next i
    If g = tt Then '' построение первоначального построения в финальной итерации цикла ''
        For i = 2 To n - 1
            X(i) = X(i - 1) + dx
            C(i) = f(X(i))
        Next i
        For i = 1 To n
            Picture1.PSet (X(i) / Xmax * Picture1.ScaleWidth, Picture1.ScaleHeight - C(i) / 2E+20 * Picture1.ScaleHeight)
        Next i
    End If
Next g
End Sub
