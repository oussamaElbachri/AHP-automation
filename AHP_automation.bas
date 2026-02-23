Attribute VB_Name = "Module6"
Sub AHP()

'clearing content
ActiveSheet.Cells.ClearContents



Dim comprange1 As Range
    Dim comprange2 As Range
    Dim comprange3 As Range
    Dim cell As Range



'taking criterions and alternatives names

Dim i As Integer
Dim j As Integer
Dim Criterion  As String
Dim Alternative As String

Criterion = "a"
Alternative = "a"

Do While Criterion <> ""

    Criterion = InputBox("enter Criterion number " & i, "enter criterion")
    Range("A1").Offset(i, 0).Value = Criterion
    i = i + 1

Loop


Do While Alternative <> ""

    Alternative = InputBox("enter alternative number " & j, "enter alternative")
    Range("B1").Offset(j, 0).Value = Alternative
    j = j + 1

Loop

'''tables Creation
'Criterions table

Dim currentrangeC As Range
Dim TrangeC As Range

Set currentrangeC = Range("A1")
Set TrangeC = currentrangeC.Resize(i - 1, 1)

TrangeC.Select
Selection.Copy
Range("E1").Select
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
Application.CutCopyMode = False

TrangeC.Select
Selection.Copy
Range("D2").Select
ActiveSheet.Paste
Application.CutCopyMode = False

'''diagonal felling
Dim m As Integer
Dim p As Integer
p = 1
m = 1

Range("E1").Offset(1, 0).Select

For m = 1 To i - 1
    ActiveCell.Value = 1
    ActiveCell.Offset(1, 1).Select
Next m

'''taking comparison points for criterions
Dim variable1 As Single

Range("E1").Select
For p = 1 To i - 2
    ActiveCell.Offset(1, 0).Select
    For m = 1 To i - 1 - p
        variable1 = Application.Evaluate(InputBox("enter comparison point between " & ActiveCell.Offset(0, -1).Value & " and " & ActiveCell.Offset(-p, m + p - 1).Value & "(on a scale of 1 to 9)"))
        ActiveCell.Offset(0, m + p - 1).Value = variable1
        ActiveCell.Offset(m, p - 1).Value = 1 / variable1
    Next m
Next p

'sums of rows
Dim sumrange As Range
Dim weight As Single
Range("E1").Select
Dim name1 As String
Dim name2 As String
Dim name3 As String

name1 = "Sum of rows"
name2 = "Weights"
name3 = "Sums :"

ActiveCell.Offset(0, i - 1).Value = name1
ActiveCell.Offset(0, i).Value = name2
ActiveCell.Offset(i, i - 2).Value = name3


For p = 1 To i - 1
    Set sumrange = Range("E1").Offset(p, 0).Resize(1, i - 1)
    ActiveCell.Offset(p, i - 1).Value = Application.WorksheetFunction.Sum(sumrange)
Next p

'total of sums
Set sumrange = Range("E1").Offset(1, i - 1).Resize(i - 1, 1)
ActiveCell.Offset(i, i - 1).Value = Application.WorksheetFunction.Sum(sumrange)

'weights
For p = 1 To i - 1
    weight = Range("E1").Offset(p, i - 1).Value / ActiveCell.Offset(i, i - 1).Value
    ActiveCell.Offset(p, i).Value = weight
Next p

'sum of weights
Set sumrange = Range("E1").Offset(1, i).Resize(i - 1, 1)
ActiveCell.Offset(i, i).Value = Application.WorksheetFunction.Sum(sumrange)

'-------------------------------------------------------------------------------------------
'matrix squaring
'creating the table
Dim currentrangeCm As Range
Dim TrangeCm As Range
Dim matrixA As Variant
Dim matrixB As Variant
Dim precision As Single
Dim t As Integer
t = 1
precision = 2

Do While precision = 2

    Set currentrangeCm = Range("A1")
    Set TrangeCm = currentrangeCm.Resize(i - 1, 1)
    
    TrangeCm.Select
    Selection.Copy
    Range("E1").Offset(t * (i + 2), 0).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Application.CutCopyMode = False
    
    TrangeCm.Select
    Selection.Copy
    Range("D2").Offset(t * (i + 2), 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'squaring the matrix
    matrixA = Range("E1").Offset(1 + (t - 1) * (i + 2), 0).Resize(i - 1, i - 1)
    matrixB = Application.WorksheetFunction.MMult(matrixA, matrixA)
    Range("E1").Offset(t * (i + 3) - t + 1, 0).Resize(i - 1, i - 1) = matrixB
    
    ''sums of rows
    Range("E1").Select
    ActiveCell.Offset(t * (i + 2), i - 1).Value = name1
    ActiveCell.Offset(t * (i + 2), i).Value = name2
    ActiveCell.Offset(i + t * (i + 2), i - 2).Value = name3
    
    
    For p = 1 To i - 1
        Set sumrange = Range("E1").Offset(p + t * (i + 2), 0).Resize(1, i - 1)
        ActiveCell.Offset(p + t * (i + 2), i - 1).Value = Application.WorksheetFunction.Sum(sumrange)
    
    Next p
    
    'total of sums
    Set sumrange = Range("E1").Offset(1 + t * (i + 2), i - 1).Resize(i - 1, 1)
    ActiveCell.Offset(i + t * (i + 2), i - 1).Value = Application.WorksheetFunction.Sum(sumrange)
    
    'weights
    For p = 1 To i - 1
        weight = Range("E1").Offset(p + t * (i + 2), i - 1).Value / ActiveCell.Offset(i + t * (i + 2), i - 1).Value
        ActiveCell.Offset(p + t * (i + 2), i).Value = weight
    Next p
    
    'sum of weights
    Set sumrange = Range("E1").Offset(1 + t * (i + 2), i).Resize(i - 1, 1)
    ActiveCell.Offset(i + t * (i + 2), i).Value = Application.WorksheetFunction.Sum(sumrange)
    
    
    
    If t > 1 Then
        Set comprange1 = ActiveCell.Offset(1 + t * (i + 2), i).Resize(i - 1, 1)
        Set comprange2 = ActiveCell.Offset(1 + (t - 1) * (i + 2), i).Resize(i - 1, 1)
        Set comprange3 = ActiveCell.Offset(1 + (t + 1) * (i + 2), i).Resize(i - 1, 1)
        comprange3.Value = comprange1.Value
        comprange2.Copy
        comprange3.PasteSpecial Paste:=xlPasteValues, Operation:=xlSubtract
        Application.CutCopyMode = False
        Dim u As Boolean
        u = False
        For Each cell In comprange3
         If Abs(cell.Value) > 0.01 Then
            u = True
         End If
        Next cell
        
        If u = False Then
            comprange3.ClearContents
            Exit Do
        End If
        
        comprange3.ClearContents
    End If
    t = t + 1
    
Loop

Dim l As Integer
l = t
'-------------------------------------------------------------------------------------------

'Alternatives tables

Dim currentrangea As Range
Dim Trangea As Range
Dim o As Integer

Set currentrangea = Range("B1")
Set Trangea = currentrangea.Resize(j - 1, 1)

For o = 0 To i - 2

    Trangea.Select
    Selection.Copy
    Range("E1").Offset(o * (j + 2), i + 3).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Application.CutCopyMode = False
    
    Trangea.Select
    Selection.Copy
    Range("D2").Offset(o * (j + 2), i + 3).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Dim n As Integer
    n = 1
    
    Range("E1").Offset(o * (j + 2) + 1, i + 3).Select
    
    '''diagonal felling
    For n = 1 To j - 1
        ActiveCell.Value = 1
        ActiveCell.Offset(1, 1).Select
    Next n
    
    '''taking comparison points for alternatives
    Range("E1").Offset(o * (j + 2), i + 3).Select
    Dim variable2 As Single
    
    For p = 1 To j - 2
        ActiveCell.Offset(1, 0).Select
        For m = 1 To j - 1 - p
            variable2 = Application.Evaluate(InputBox("enter comparison point between " & ActiveCell.Offset(0, -1).Value & " and " & ActiveCell.Offset(-p, m + p - 1).Value & " for criterion " & Range("A1").Offset(o, 0).Value & "(on a scale of 1 to 9)"))
            ActiveCell.Offset(0, m + p - 1).Value = variable2
            ActiveCell.Offset(m, p - 1).Value = 1 / variable2
        Next m
    Next p
    
    ActiveCell.Offset(-j + 2, -1).Value = Range("A1").Offset(o, 0).Value
    
    'sums of rows
    Dim sumrangea As Range
    Dim weighta As Single
    ActiveCell.Offset(-j + 2, -1).Select
    Dim name1a As String
    Dim name2a As String
    Dim name3a As String
    
    name1a = "Sum of rows"
    name2a = "Scores"
    name3a = "Sums :"
    
    ActiveCell.Offset(0, j).Value = name1a
    ActiveCell.Offset(0, j + 1).Value = name2a
    ActiveCell.Offset(j, j - 1).Value = name3a
    
    
    For p = 1 To j - 1
        Set sumrangea = ActiveCell.Offset(p, 1).Resize(1, j - 1)
        ActiveCell.Offset(p, j).Value = Application.WorksheetFunction.Sum(sumrangea)
    Next p
    
    'total of sums
    Set sumrange = ActiveCell.Offset(1, j).Resize(j - 1, 1)
    ActiveCell.Offset(j, j).Value = Application.WorksheetFunction.Sum(sumrange)
    'weights
    For p = 1 To j - 1
        weight = ActiveCell.Offset(p, j).Value / ActiveCell.Offset(j, j).Value
        ActiveCell.Offset(p, j + 1).Value = weight
    Next p
    
    'sum of weights
    Set sumrange = ActiveCell.Offset(1, j + 1).Resize(j - 1, 1)
    ActiveCell.Offset(j, j + 1).Value = Application.WorksheetFunction.Sum(sumrange)
    
    
    '----------------------------------------------
    'squaring alternatives table
     Dim currentrangeam As Range
     Dim Trangeam As Range
     Dim matrixAa As Variant
     Dim matrixBa As Variant
     
        Set currentrangeam = Range("B1")
        Set Trangeam = currentrangeam.Resize(j - 1, 1)
        
        Trangeam.Select
        Selection.Copy
        Range("E1").Offset(o * (j + 2), i + 3 + j + 3).Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Application.CutCopyMode = False
        
        Trangeam.Select
        Selection.Copy
        Range("D2").Offset(o * (j + 2), i + 3 + j + 3).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        'squaring the matrix
        matrixAa = Range("E1").Offset(o * (j + 2) + 1, i + 3).Resize(j - 1, j - 1)
        matrixBa = Application.WorksheetFunction.MMult(matrixAa, matrixAa)
        Range("E1").Offset(o * (j + 3) - o + 1, i + 3 + j + 3).Resize(j - 1, j - 1) = matrixBa

        
        ''sums of rows
        Range("E1").Offset(o * (j + 3) - o, i + 3 + j + 3 + j - 1).Value = name1a
        Range("E1").Offset(o * (j + 3) - o, i + 3 + j + 3 + j).Value = name2a
        Range("E1").Offset(o * (j + 3) - o + j, i + 3 + j + 3 + j - 2).Value = name3a
        
        
        For p = 1 To j - 1
            Set sumrange = Range("E1").Offset(p + o * (j + 2), i + 3 + j + 3).Resize(1, j - 1)
            Range("E1").Offset(p + o * (j + 2), i + 3 + j + 3 + j - 1).Value = Application.WorksheetFunction.Sum(sumrange)
        Next p
       
        
        'total of sums
        Set sumrange = Range("E1").Offset(1 + o * (j + 2), i + 3 + j + 3 + j - 1).Resize(j - 1, 1)
        Range("E1").Offset(o * (j + 2) + j, i + 3 + j + 3 + j - 1).Value = Application.WorksheetFunction.Sum(sumrange)

        
        'weights
        For p = 1 To j - 1
            weight = Range("E1").Offset(p + o * (j + 2), i + 3 + j + 3 + j - 1).Value / Range("E1").Offset(o * (j + 2) + j, i + 3 + j + 3 + j - 1).Value
            Range("E1").Offset(p + o * (j + 2), i + 3 + j + 3 + j).Value = weight
        Next p
        
        'sum of weights
        Set sumrange = Range("E1").Offset(o * (j + 2), i + 3 + j + 3 + j).Resize(j, 1)
        Range("E1").Offset(o * (j + 2) + j, i + 3 + j + 3 + j).Value = Application.WorksheetFunction.Sum(sumrange)
    
    '----------------------------------------------

    
Next o

'Range("E1").Offset(i * (j + 2), i + j + 4)
Dim weightsrange As Range
Dim scoresrange As Range
Dim placeholder As Range
n = 0
Set weightsrange = Range("E1").Offset(t * (i + 2) + 1, i).Resize(i - 1, 1)
Dim name4 As String
name4 = "total score"
Dim f As Integer
f = 0
weightsrange.Select

For p = 1 To j - 1
    Set placeholder = Range("E1").Offset(3 + (p - 1) * (j + 2), i + 2 * j + 9).Resize(i - 1, 1)
    Set scoresrange = Range("E1").Offset(3 + (p - 1) * (j + 2), i + 2 * j + 8).Resize(i - 1, 1)
    

    For Each cell In placeholder
        Range("E1").Offset(f + (j + 2) * n + 1, i + 2 * j + 6).Select
        Selection.Copy
        cell.Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        n = n + 1
    Next cell
    n = 0
    f = f + 1
    
    weightsrange.Select
    Selection.Copy
    scoresrange.Select
    ActiveSheet.Paste
    placeholder.Select
    Selection.Copy
    scoresrange.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply
        Application.CutCopyMode = False
    Range("E1").Offset(2 + i + (p - 1) * (j + 2), i + 2 * j + 9).Value = Application.WorksheetFunction.Sum(scoresrange)
    Range("E1").Offset(2 + i + (p - 1) * (j + 2), i + 2 * j + 8).Value = name4
    placeholder.ClearContents
Next p

Dim g As Integer
Dim end_resault As Range
Dim name5 As String
g = 1
Set end_resault = Range("E1").Offset(i + 5, i + 2 * j + 12).Resize(j - 1, 1)

end_resault.Select



name5 = "AHP end resaults"
Range("E1").Offset(1 + i + j + 2, i + 2 * j + 14).Value = name5

For Each cell In end_resault
    Range("E1").Offset(2 + i + (g - 1) * (j + 2), i + 2 * j + 9).Select
    Selection.Copy
    cell.Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    g = g + 1
Next cell

p = 1

For Each cell In Trangea
    cell.Select
    Selection.Copy
    Range("E1").Offset(i + 5 + p - 1, i + 2 * j + 11).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    p = p + 1
Next cell

Dim maxr As Range
p = 1
Dim maxv As Variant
Set maxr = Range("E1").Offset(i + 5, i + 2 * j + 12).Resize(j - 1, 1)

For Each cell In maxr

    If cell.Value = Application.WorksheetFunction.Max(Range("E1").Offset(i + 5, i + 2 * j + 12).Resize(j - 1, 1)) Then
        Range("E1").Offset(1 + i + j + 2, i + 2 * j + 15).Value = Range("E1").Offset(i + 5 + p - 1, i + 2 * j + 11).Value
    End If
    p = p + 1

Next cell

MsgBox ("the best choise from all the alternatives based on the AHP methode is " & Range("E1").Offset(1 + i + j + 2, i + 2 * j + 15).Value & " with the score of " & Application.WorksheetFunction.Max(Range("E1").Offset(i + 5, i + 2 * j + 12).Resize(j - 1, 1)))



Dim ws As Worksheet
Dim rngData As Range
Dim rngData2 As Range
Dim rngData3 As Range
Dim rngData4 As Range
Dim chartObj As ChartObject
Dim chartObj2 As ChartObject
Dim rangeBack As Range

Set rangeBack = Range("E1").Offset(1 + (l - 1) * (i + 2), i).Resize(i - 1, 1)
Set rngData2 = Range("E1").Offset(i + 5 + j, i + 2 * j + 21).Resize(i - 1, 1)
rangeBack.Select
Selection.Copy
rngData2.Select
ActiveSheet.Paste
Application.CutCopyMode = False


Set rngData3 = Range("E1").Offset(i + 5 + j, i + 2 * j + 20).Resize(i - 1, 1)
Range("A1").Resize(i - 1, 1).Select
Selection.Copy
rngData3.Select
ActiveSheet.Paste
Application.CutCopyMode = False



Set rngData4 = Range("E1").Offset(i + 5 + j, i + 2 * j + 20).Resize(i - 1, 2)

Sheets.Add.Name = "Sheet2"



Set ws = ThisWorkbook.Sheets("Sheet2")
Set rngData = Range("E1").Offset(i + 5, i + 2 * j + 11).Resize(j - 1, 2)
    
   
Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("B4").Left, _
        Top:=ws.Range("B4").Top, _
        Width:=400, _
        Height:=300)
        

With chartObj.Chart
    .SetSourceData Source:=rngData
    .ChartType = xlColumnClustered
    .HasTitle = True
    .ChartTitle.Text = "total scores for alternatives"
    .HasLegend = True
End With

Set chartObj2 = ws.ChartObjects.Add( _
        Left:=ws.Range("M4").Left, _
        Top:=ws.Range("M4").Top, _
        Width:=400, _
        Height:=300)
        

With chartObj2.Chart
    .SetSourceData Source:=rngData4
    .ChartType = xlColumnClustered
    .HasTitle = True
    .ChartTitle.Text = "Criterion wiegths"
    .HasLegend = True
    .ChartColor = 12
End With


End Sub


