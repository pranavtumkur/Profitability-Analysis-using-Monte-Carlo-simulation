VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Monte Carlo Simulation"
   ClientHeight    =   6720
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   8412.001
   OleObjectBlob   =   "Code behind User form to trigger simulation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub CommandButton1_Click()
Dim datamin As Double, datamax As Double, datarange As Double
Dim lowbins As Integer, highbins As Integer, nbins As Integer, hist As Integer
Dim binrangeinit As Double, binrangefinal As Double
Dim bins() As Double, bincenters() As Double, j As Integer
Dim c As Integer, i As Integer, R() As Double
Dim bincounts() As Integer, ChartRange As String, nr As String, o As Integer
Dim countnpv As Integer
Dim satu(1000), dua(1000), tiga(1000), empat(1000), lima(1000), enam(1000), tujuh(1000), delapan(1000), sembilan(1000)

Dim tWB As Workbook
Set tWB = ThisWorkbook
tWB.Activate

ReDim R(nsimulations) As Double
On Error Resume Next
Application.DisplayAlerts = False
Sheets("Main").Select
For i = 1 To nsimulations
    
    Range("B3") = discretecland(TextBox3, TextBox2, TextBox1, TextBox13, TextBox12, TextBox11)
    
    Range("B4") = betapertcroy(TextBox8, TextBox9, TextBox10)

    Range("B5") = TDC(-1 * TextBox14, TextBox15)

    Range("B6") = -1 * ((-1 * TextBox16) + ((-1 * TextBox17) - (-1 * TextBox16)) * Rnd)

    Range("B7") = -WorksheetFunction.Norm_Inv(Rnd, -1 * TextBox18, TextBox19)

    Range("E3") = salesrev(TextBox22, TextBox20, TextBox21)

    Range("H3") = triangular(Rnd, TextBox25, TextBox23, TextBox24)

    Range("E4") = discretetax(TextBox26, TextBox27, TextBox28, TextBox29)

    Range("H4") = TextBox30 + (TextBox31 - TextBox30) * Rnd

    R(i) = Range("N24")
Next i


Sheets("Sheet1").Range("A1:A100") = WorksheetFunction.Transpose(R)
datamin = WorksheetFunction.min(R)
datamax = WorksheetFunction.Max(R)
datarange = datamax - datamin
lowbins = Int(WorksheetFunction.Log(nsimulations, 2)) + 1
highbins = Int(Sqr(nsimulations))
nbins = (lowbins + highbins) / 2
binrangeinit = datarange / nbins
ReDim bins(1) As Double
If binrangeinit < 1 Then
     c = 1
     Do
        If 10 * binrangeinit > 1 Then
            binrangefinal = 10 * binrangeinit Mod 10
            Exit Do
        Else
            binrangeinit = 10 * binrangeinit
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal / 10 ^ c
ElseIf binrangeinit < 10 Then
    binrangefinal = binrangeinit Mod 10
Else
    c = 1
    Do
        If binrangeinit / 10 < 10 Then
            binrangefinal = binrangeinit / 10 Mod 10
            Exit Do
        Else
            binrangeinit = binrangeinit / 10
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal * 10 ^ c
End If
i = 1
bins(1) = (datamin - ((datamin) - (binrangefinal * Fix(datamin / binrangefinal))))
Do
    i = i + 1
    ReDim Preserve bins(i) As Double
    bins(i) = bins(i - 1) + binrangefinal
Loop Until bins(i) > datamax
nbins = i
ReDim Preserve bincounts(nbins - 1) As Integer
ReDim Preserve bincenters(nbins - 1) As Double
For j = 1 To nbins - 1
    c = 0
    For i = 1 To nsimulations
        If R(i) > bins(j) And R(i) <= bins(j + 1) Then
            c = c + 1
        End If
    Next i
    bincounts(j) = c
    bincenters(j) = (bins(j) + bins(j + 1)) / 2
Next j
Sheets("Histogram").Cells.ClearContents
Sheets("Histogram Data").Select
Cells.Clear
Range("A1").Select
Range("A1:A" & nbins - 1) = WorksheetFunction.Transpose(bincenters)
Range("B1:B" & nbins - 1) = WorksheetFunction.Transpose(bincounts)
UserForm1.Hide
Application.ScreenUpdating = False
Charts("Histogram").Delete
ActiveCell.Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    nr = Selection.Rows.Count
    ChartRange = Selection.Addres
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Histogram Data'!" & ChartRange)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Delete
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).XValues = "='Histogram Data'!" & "$A$1:$A$" & nbins - 1
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Caption = "Count"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "Bin Center"
    ActiveChart.ChartArea.Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Histogram"
For o = 1 To UserForm1.nsimulations
    If R(o) > 0 Then
        countnpv = countnpv + 1
    End If
Next o
Application.ScreenUpdating = False
MsgBox ((countnpv / UserForm1.nsimulations) * 100 & "% of the simulations were found profitable")

hist = MsgBox("Do you want to view a histogram of the simulation results?", vbYesNo)
If hist = 6 Then
    Sheets("Histogram").Activate
Else
    Sheets("Histogram").Visible = False
    Sheets("Main").Activate
End If

End Sub

Function discretecland(prb1 As Double, prb2 As Double, prb3 As Double, value1 As Double, value2 As Double, value3 As Double) As Double
Dim Rnumber As Double
Rnumber = Rnd * 100
If Rnumber < prb1 Then
    discretecland = value1
ElseIf Rnumber < prb2 + prb1 Then
    discretecland = value2
Else
    discretecland = value3
End If
End Function

Function betapertcroy(TextBox8 As Double, TextBox9 As Double, TextBox10 As Double) As Double
Dim alpha As Double, beta As Double, valbetinv As Double, c As Double, a As Double, b As Double

a = TextBox8
c = (TextBox10)
b = TextBox9
alpha = ((4 * b + c - 5 * a) / (c - a))
beta = ((5 * c - a - 4 * b) / (c - a))
betapertcroy = -WorksheetFunction.Beta_Inv(Rnd, alpha, beta, -1 * TextBox8, -1 * TextBox10)

End Function


Function TDC(ave As Double, std As Double) As Double
TDC = -WorksheetFunction.Norm_Inv(Rnd, ave, std)

End Function

Function salesrev(TextBox8 As Double, TextBox9 As Double, TextBox10 As Double) As Double
Dim alpha As Double, beta As Double, valbetinv As Double, c As Double, a As Double, b As Double

a = TextBox8
c = (TextBox10)
b = TextBox9
alpha = ((4 * b + c - 5 * a) / (c - a))
beta = ((5 * c - a - 4 * b) / (c - a))
valbetinv = WorksheetFunction.Beta_Inv(Rnd, alpha, beta, a, c)
salesrev = (valbetinv)

End Function

Function triangular(Rnd, L As Double, M As Double, U As Double) As Double
Dim P As Double, triangular1 As Double
P = Rnd
L = -1 * L
M = -1 * M
U = -1 * U

Dim a As Double, b As Double, c As Double
If P < (M - L) / (U - L) Then
    a = 1
    b = -2 * L
    c = L ^ 2 - P * (M - L) * (U - L)
    triangular1 = (-b + Sqr(b ^ 2 - 4 * a * c)) / 2 / a
ElseIf P <= 1 Then
    a = 1
    b = -2 * U
    c = U ^ 2 - (1 - P) * (U - L) * (U - M)
    triangular1 = (-b - Sqr(b ^ 2 - 4 * a * c)) / 2 / a
End If


triangular = -1 * triangular1

End Function

Function discretetax(prb1 As Double, prb2 As Double, value1 As Double, value2 As Double) As Double
Dim Rnumber As Double
Rnumber = Rnd * 100
If Rnumber < prb1 Then
    discretetax = value1
Else
    discretetax = value2
End If
End Function

Private Sub CommandButton2_Click()
Unload UserForm1
End Sub
