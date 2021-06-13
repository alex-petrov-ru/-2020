VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23970
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   23970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   13320
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   22800
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   13320
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   10320
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H8000000E&
      Height          =   6735
      Left            =   13320
      ScaleHeight     =   6675
      ScaleWidth      =   10155
      TabIndex        =   2
      Top             =   1320
      Width           =   10215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "СМОДЕЛИРОВАТЬ"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   9000
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000E&
      Height          =   6735
      Left            =   1320
      ScaleHeight     =   6675
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Label Label4 
      Caption         =   "S"
      Height          =   255
      Left            =   14160
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "S"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "мкм"
      Height          =   255
      Left            =   23520
      TabIndex        =   9
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "мкм"
      Height          =   255
      Left            =   11160
      TabIndex        =   8
      Top             =   8160
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x(), G(), xm(), Gn(), Gp(), dpn(), dnp(), dpn1(), dnp1(), a(), d1(), b(), r(), delta(), lamda(), lx(), sl(), sjn(), sjw(), sjp(), Uvn() As Single
Private Sub Command1_Click()

Nd = 4E+16                    'концентрация в n области
Na = 100 * Nd                'концентрация в p+ области
xj = 0.45 * 0.0001            'глубина перехода см
tau = 0.0000001                  'время жизни
first_x = 0.4
Eg = 1.424 + 1.247 * first_x                     'ширина запрещенной зоны
eps = 12.9 - 2.84 * first_x                    'диэл. константа
Nc = 2.5E+19 * (0.063 + 0.083 * first_x) ^ 1.5 'эффективная плотность состояний в зоне проводимости
Nv = 2.5E+19 * (0.51 + 0.25 * first_x) ^ 1.5 'эффективная плотность состояний в валентной зоне

ni = Exp(-Eg / (2 * 0.026)) * (Nc * Nv) ^ 0.5          'собственная концентрация
mu_n = -255 + 1160 * first_x - 720 * first_x ^ 2         'подвижность электронов
mu_p = 370 - 970 * first_x + 740 * first_x ^ 2           'подвижность дырок
Dp = 0.026 * mu_p                       'коэф. диффузии дырок
Dn = 0.026 * mu_n                       'коэф. диффузии электронов
S = 50000                    'скорость поверхностной рекомбинации

Q = 1.6E-19                   'элементарный заряд
eps0 = 8.85E-14               'диэл. константа
I0 = 1E+15                    'интенсивность излучения

d = 5 * 0.0001                'толщина структуры в см
Uv = 0                        'внешнее смещение


'Оптические характеристики материала
lambda = 1.24 / (Eg)
En = 1.24 / lambda
n = 3.28                        'коэф. преломления
re = (n - 1) ^ 2 / (n + 1) ^ 2 'коэф. отражения
alpha = 405.2


If En < Eg Then alpha = 0      'п/п прозрачен, если энергия фотонов меньше Eg
m = 10000                      'кол-во шагов по глубине
dx = d / (m - 1)

ReDim x(m), G(m), xm(m)

'Распределение неосновных н.з.
fik = 0.026 * Log(Na * Nd / (ni ^ 2))
Nb = Na * Nd / (Na + Nd)
Wp = ((2 * eps * eps0 * Nb * (fik - Uv) / Q) ^ 0.5) / Na
Wn = ((2 * eps * eps0 * Nb * (fik - Uv) / Q) ^ 0.5) / Nd

W = Wn + Wp

sp = CInt((xj - Wp) * m / d)        'кол-во индексов отделенных для p-области
sn = CInt((d - xj - Wn) * m / d)    'кол-во индексов отделенных для n-области
sw = m - sn - sp

ReDim Gn(sn), Gp(sp), dpn(sn), dnp(sp), dpn1(sn), dnp1(sp)
'Gn массив генерации в n области, dpn распределение избыточной конц. дырок в n области,
'dnp распределение электронов в p+ области

For i = 0 To sp - 1
Gp(i) = G(i)
Next i

For i = 0 To sn - 1
Gn(i) = G(i + sp + sw)
Next i

'Рассчитаем токи
'в p-области ток неосновных электронов
jn = Q * Dn * (dnp(sp - 2) - dnp(sp - 1)) / dx
'в n-области ток неосновных дырок
jp = Q * Dp * (dpn(1) - dpn(0)) / dx
'ток в ОПЗ
jw = Q * I0 * (1 - re) * Exp(-alpha * (xj - Wp)) - Q * I0 * (1 - re) * Exp(-alpha * (xj + Wn))  'считаем дельту между крайними точками около ОПЗ

jph = jp + jn + jw 'полный фототок

'спектральная чувствительность
P = I0 * Q * 1.24 / lambda
S = Abs(jph / P)


lambda1 = 0.19        'коротковолновая
lambda2 = 0.63        'длинноволновая

'количество шагов по длине волны
v = 500
m = 500
dlambda = (lambda2 - lambda1) / (v - 1)    'шаг
ReDim lx(v), sl(v), sjn(v), sjw(v), sjp(v)

lx(0) = lambda1    'перезадается лямбда
For i = 1 To v - 1
lx(i) = lx(i - 1) + dlambda
Next i

For ii = 0 To v - 1
'Оптические характеристики
lambda = lx(ii)
En = 1.24 / lambda
n = 542.57 * lambda ^ 6 - 1380.6 * lambda ^ 5 + 1063 * lambda ^ 4 + 48.352 * lambda ^ 3 - 445 * lambda ^ 2 + 202.41 * lambda - 23.818    'коэф. преломления
re = (n - 1) ^ 2 / (n + 1) ^ 2                           'коэф. отражения
If lambda < lambda1 Then
alpha = 1 * 10 ^ 13 * lambda ^ 6 - 2 * 10 ^ 13 * lambda ^ 5 + 2 * 10 ^ 13 * lambda ^ 4 - 6 * 10 ^ 12 * lambda ^ 3 + 1 * 10 ^ 12 * lambda ^ 2 - 1 * 10 ^ 11 * lambda - 7 * 10 ^ 9
Else
alpha = 1 * 10 ^ 9 * Exp(-29.64 * lambda)
End If

If En <= Eg Then alpha = 0

dx = d / (m - 1)

'Распределение генерации фотодиода
For i = 1 To m - 1
x(i) = x(i - 1) + dx
Next i

For i = 1 To m - 1
G(i) = I0 * (1 - re) * alpha * Exp(-x(i) * alpha)
Next i

Max = 1
For i = 0 To m - 1
If G(i) > Max Then
Max = G(i)
End If
Next i

'Выведем распределение неосновных нз
fik = 0.026 * Log(Na * Nd / (ni ^ 2))
Nb = Na * Nd / (Na + Nd)
Wp = ((2 * eps * eps0 * Nb * (fik - Uv) / Q) ^ 0.5) / Na
Wn = ((2 * eps * eps0 * Nb * (fik - Uv) / Q) ^ 0.5) / Nd

W = Wn + Wp

sp = CInt((xj - Wp) * m / d)         'кол-во индексов для p области
sn = CInt((d - xj - Wn) * m / d)    'кол-во индексов для n области
sw = m - sn - sp

For i = 0 To sp - 1
Gp(i) = G(i)
Next i

For i = 0 To sn - 1
Gn(i) = G(i + sp + sw)
Next i

ReDim a(sp), d1(sp), b(sp), r(sp), delta(sp), lamda(sp)

'Запишем начальное распредление
For i = 0 To sn - 1
dpn(i) = Gn(i) * tau
Next i

For i = 0 To sp - 1
dnp(i) = Gp(i) * tau
Next i

'Прогонка для p+ области

'граничные условия для поверхности
a(0) = (S * dx / Dn) - 1
b(0) = 1
d1(0) = 0
r(0) = 0
'граничные условия для ОПЗ
a(sp - 1) = 1
b(sp - 1) = 0
d1(sp - 1) = 0
r(sp - 1) = 0

delta(0) = -d1(0) / a(0)
lamda(0) = r(0) / a(0)

For j = 0 To 1000

For i = 1 To sp - 2
d1(i) = 1
a(i) = -2 - (dx * dx / (Dn * tau))
b(i) = 1
r(i) = -dx * dx * Gp(i) / Dn
Next i

For i = 1 To sp - 1
delta(i) = -(d1(i) / (a(i) + b(i) * delta(i - 1)))
lamda(i) = (r(i) - b(i) * lamda(i - 1)) / (a(i) + b(i) * delta(i - 1))
Next i

dnp(sp - 1) = lamda(sp - 1)

For i = sp - 2 To 0 Step -1
dnp(i) = delta(i) * dnp(i + 1) + lamda(i)
Next i

Sum = 0
For i = 0 To sp - 1
Sum = Sum + Abs(dnp(i) - dnp1(i))
Next i

dnp1() = dnp()

If Sum < 0.01 Then Exit For   'Прогонем до стационарного изменения
Next j

'n область

ReDim a(sn), d1(sn), b(sn), r(sn), delta(sn), lamda(sn)
'граничные условия для поверхности
a(0) = 1
b(0) = 0
d1(0) = 0
r(0) = 0
'граничные условия для ОПЗ
a(sn - 1) = -((S * dx / Dp) + 1)
b(sn - 1) = 1
d1(sn - 1) = 0
r(sn - 1) = 0

delta(0) = -d1(0) / a(0)
lamda(0) = r(0) / a(0)

For j = 0 To 1000
For i = 1 To sn - 2
d1(i) = 1
a(i) = -2 - (dx * dx / (Dp * tau))
b(i) = 1
r(i) = -dx * dx * Gn(i) / Dp
Next i

For i = 1 To sn - 1
delta(i) = -(d1(i) / (a(i) + b(i) * delta(i - 1)))
lamda(i) = (r(i) - b(i) * lamda(i - 1)) / (a(i) + b(i) * delta(i - 1))
Next i

dpn(sn - 1) = lamda(sn - 1)

For i = sn - 2 To 0 Step -1
dpn(i) = delta(i) * dpn(i + 1) + lamda(i)
Next i

Sum = 0
For i = 0 To sn - 1
Sum = Sum + Abs(dpn(i) - dpn1(i))
Next i

dpn1() = dpn()

If Sum < 0.01 Then Exit For  'Прогоняем до стационарного изменения
Next j

'Рассчитаем токи
'в p-области ток неосновных электронов
jn = Q * Dn * (dnp(sp - 2) - dnp(sp - 1)) / dx
'в n области ток неосновных дырок
jp = Q * Dp * (dpn(1) - dpn(0)) / dx
'ток в ОПЗ
jw = Q * I0 * (1 - re) * Exp(-alpha * (xj - Wp)) - Q * I0 * (1 - re) * Exp(-alpha * (xj + Wn))  'считаем дельту между крайними точками около ОПЗ

jph = jp + jn + jw 'полный фототок

'Спектральная чувствительность
P = I0 * Q * 1.24 / lambda
sjp(ii) = Abs(jp / P)
sjw(ii) = Abs(jw / P)
sjn(ii) = Abs(jn / P)

sl(ii) = Abs(jph / P)

Next ii

smax = 0
For i = 0 To v - 1
If sl(i) > smax Then
smax = sl(i)
Text2 = smax
End If
Next i

For i = 0 To v - 1
Picture2.PSet ((lx(i) - lambda1) * Picture2.ScaleWidth / (dlambda * (v - 1)), Picture2.ScaleHeight - sl(i) * Picture2.ScaleHeight / smax)
Picture2.PSet ((lx(i) - lambda1) * Picture2.ScaleWidth / (dlambda * (v - 1)), Picture2.ScaleHeight - sjp(i) * Picture2.ScaleHeight / smax), vbRed
Picture2.PSet ((lx(i) - lambda1) * Picture2.ScaleWidth / (dlambda * (v - 1)), Picture2.ScaleHeight - sjw(i) * Picture2.ScaleHeight / smax), vbBlue
Picture2.PSet ((lx(i) - lambda1) * Picture2.ScaleWidth / (dlambda * (v - 1)), Picture2.ScaleHeight - sjn(i) * Picture2.ScaleHeight / smax), QBColor(2)
Next i

U1 = -1
U2 = -5

k = 4    'для разных кривых по цветам
dU = (Abs(U2) - Abs(U1)) / (k - 1)

ReDim Uvn(k)

Uvn(0) = 0
For i = 1 To k - 1
Uvn(i) = Uvn(i - 1) - dU
Next i

For z = 0 To k - 1

Uv = Uvn(z) 'старая переменная переходит в новую (внешнее смещение)

v = 500
m = 500

dlambda = (lambda2 - lambda1) / (v - 1)

ReDim lx(v), sl(v), sjn(v), sjw(v), sjp(v)

lx(0) = lambda1
For i = 1 To v - 1
lx(i) = lx(i - 1) + dlambda
Next i

For ii = 0 To v - 1

'Оптические характеристики
lambda = lx(ii)
En = 1.24 / lambda
n = 542.57 * lambda ^ 6 - 1380.6 * lambda ^ 5 + 1063 * lambda ^ 4 + 48.352 * lambda ^ 3 - 445 * lambda ^ 2 + 202.41 * lambda - 23.818 'коэф. преломления
re = (n - 1) ^ 2 / (n + 1) ^ 2 'коэф. отражения
If lambda < lambda1 Then
alpha = 1 * 10 ^ 13 * lambda ^ 6 - 2 * 10 ^ 13 * lambda ^ 5 + 2 * 10 ^ 13 * lambda ^ 4 - 6 * 10 ^ 12 * lambda ^ 3 + 1 * 10 ^ 12 * lambda ^ 2 - 1 * 10 ^ 11 * lambda - 7 * 10 ^ 9
Else
alpha = 1 * 10 ^ 9 * Exp(-29.64 * lambda)
End If
If En <= Eg Then alpha = 0

dx = d / (m - 1)

ReDim x(m), G(m), xm(m)

'Распределение генерации фотодиода
For i = 1 To m - 1
x(i) = x(i - 1) + dx
Next i

For i = 1 To m - 1
G(i) = I0 * (1 - re) * alpha * Exp(-x(i) * alpha)
Next i

Max = 1
For i = 0 To m - 1
If G(i) > Max Then
Max = G(i)
End If
Next i

'Выведем распределение неосновных н.з.
fik = 0.026 * Log(Na * Nd / (ni ^ 2))
Nb = Na * Nd / (Na + Nd)
Wp = ((2 * eps * eps0 * Nb * (fik - Uv) / Q) ^ 0.5) / Na
Wn = ((2 * eps * eps0 * Nb * (fik - Uv) / Q) ^ 0.5) / Nd

W = Wn + Wp

sp = CInt((xj - Wp) * m / d)         'кол-во индексов отделенных для p-области
sn = CInt((d - xj - Wn) * m / d)    'кол-во индексов отделенных для n-области
sw = m - sn - sp

ReDim Gn(sn), Gp(sp), dpn(sn), dnp(sp), dpn1(sn), dnp1(sp)
 'Gn массив генерации в n области, dpn распределение избыточной конц. дырок в n-области,
'dnp распределение электронов в p+ области

For i = 0 To sp - 1
Gp(i) = G(i)
Next i

For i = 0 To sn - 1
Gn(i) = G(i + sp + sw)
Next i

ReDim a(sp), d1(sp), b(sp), r(sp), delta(sp), lamda(sp)

'Запишем начальное распределение
For i = 0 To sn - 1
dpn(i) = Gn(i) * tau
Next i

For i = 0 To sp - 1
dnp(i) = Gp(i) * tau
Next i

'Прогонка для p+ области

'граничные условия для поверхности
a(0) = (S * dx / Dn) - 1
b(0) = 1
d1(0) = 0
r(0) = 0
'граничные условия для ОПЗ
a(sp - 1) = 1
b(sp - 1) = 0
d1(sp - 1) = 0
r(sp - 1) = 0

delta(0) = -d1(0) / a(0)
lamda(0) = r(0) / a(0)

For j = 0 To 1000

For i = 1 To sp - 2
d1(i) = 1
a(i) = -2 - (dx * dx / (Dn * tau))
b(i) = 1
r(i) = -dx * dx * Gp(i) / Dn
Next i

For i = 1 To sp - 1
delta(i) = -(d1(i) / (a(i) + b(i) * delta(i - 1)))
lamda(i) = (r(i) - b(i) * lamda(i - 1)) / (a(i) + b(i) * delta(i - 1))
Next i

dnp(sp - 1) = lamda(sp - 1)

For i = sp - 2 To 0 Step -1
dnp(i) = delta(i) * dnp(i + 1) + lamda(i)
Next i

Sum = 0
For i = 0 To sp - 1
Sum = Sum + Abs(dnp(i) - dnp1(i))
Next i

dnp1() = dnp()

If Sum < 0.01 Then Exit For   'Прогоняем до стационарного изменения
Next j

'n область

ReDim a(sn), d1(sn), b(sn), r(sn), delta(sn), lamda(sn)
'граничные условия для поверхности
a(0) = 1
b(0) = 0
d1(0) = 0
r(0) = 0
'граничные условия для ОПЗ
a(sn - 1) = -((S * dx / Dp) + 1)
b(sn - 1) = 1
d1(sn - 1) = 0
r(sn - 1) = 0

delta(0) = -d1(0) / a(0)
lamda(0) = r(0) / a(0)

For j = 0 To 1000
For i = 1 To sn - 2
d1(i) = 1
a(i) = -2 - (dx * dx / (Dp * tau))
b(i) = 1
r(i) = -dx * dx * Gn(i) / Dp
Next i

For i = 1 To sn - 1
delta(i) = -(d1(i) / (a(i) + b(i) * delta(i - 1)))
lamda(i) = (r(i) - b(i) * lamda(i - 1)) / (a(i) + b(i) * delta(i - 1))
Next i

dpn(sn - 1) = lamda(sn - 1)

For i = sn - 2 To 0 Step -1
dpn(i) = delta(i) * dpn(i + 1) + lamda(i)
Next i

Sum = 0
For i = 0 To sn - 1
Sum = Sum + Abs(dpn(i) - dpn1(i))
Next i

dpn1() = dpn()

If Sum < 0.01 Then Exit For  'Прогонем до стационарного изменения
Next j

'Рассчитаем токи
'в p-области ток неосновных электронов
jn = Q * Dn * (dnp(sp - 2) - dnp(sp - 1)) / dx
'в n-области ток неосновных дырок
jp = Q * Dp * (dpn(1) - dpn(0)) / dx
'для ОПЗ
jw = Q * I0 * (1 - re) * Exp(-alpha * (xj - Wp)) - Q * I0 * (1 - re) * Exp(-alpha * (xj + Wn))  'считаем дельту между крайними точками около ОПЗ
jph = jp + jn + jw 'полный фототок

'Спектральная чуствительность
P = I0 * Q * 1.24 / lambda
sjw(ii) = Abs(jw / P)
sl(ii) = Abs(jph / P)

Next ii

smax = 0
For i = 0 To v - 1
If sl(i) > smax Then
smax = sl(i)
Text1 = smax
End If
Next i

If z = 0 Then
For i = 0 To v - 1

Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sjw(i) * Picture3.ScaleHeight / smax), vbRed
Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sl(i) * Picture3.ScaleHeight / smax), vbRed
Next i
End If

If z = 1 Then
For i = 0 To v - 1

Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sjw(i) * Picture3.ScaleHeight / smax), vbBlue
Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sl(i) * Picture3.ScaleHeight / smax), vbBlue
Next i
End If

If z = 2 Then
For i = 0 To v - 1

Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sjw(i) * Picture3.ScaleHeight / smax), vbGreen
Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sl(i) * Picture3.ScaleHeight / smax), vbGreen
Next i
End If

If z = 3 Then
For i = 0 To v - 1

Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sjw(i) * Picture3.ScaleHeight / smax), vbBlack
Picture3.PSet ((lx(i) - lambda1) * Picture3.ScaleWidth / (dlambda * (v - 1)), Picture3.ScaleHeight - sl(i) * Picture3.ScaleHeight / smax), vbBlack
Next i
End If

Next z
Text3 = lambda1
Text4 = lambda2
Text5 = lambda1
Text6 = lambda2
End Sub

