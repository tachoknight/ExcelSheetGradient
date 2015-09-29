VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
 
Sub Test()

    ' This example will set the block of rows from A2 to E22 from red to
    ' a reddish-white.
    GradientCellBlock "A2:E22", Array(255, 0, 0), Array(250, 240, 230)
 
End Sub
Sub GradientCellBlock(selectedBlock As String, fromColorRGB As Variant, toColorRGB As Variant)
    
    ' The range of cells we're working with
    Dim rangeBlock As Range
    
    ' Our loop counter
    Dim i As Integer
    
    ' The colors are defined as doubles to eliminate having to cast from int on
    ' every loop iteration
    Dim fromRed As Double
    Dim fromGreen As Double
    Dim fromBlue As Double
    
    Dim toRed As Double
    Dim toGreen As Double
    Dim toBlue As Double
       
    ' But we need the calculated color to be integers
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    
    Dim ratio As Double
    
    ' Get our from and to colors from the respective arrays
    fromRed = CDbl(fromColorRGB(0))
    fromGreen = CDbl(fromColorRGB(1))
    fromBlue = CDbl(fromColorRGB(2))

    toRed = CDbl(toColorRGB(0))
    toGreen = CDbl(toColorRGB(1))
    toBlue = CDbl(toColorRGB(2))
    
    ' And convert the string to a range object
    Set rangeBlock = Range(selectedBlock)
    
    ' Now for each row in the block...
    For i = 1 To rangeBlock.Rows.Count
        ' ... calculate the ratio ...
        ratio = CDbl(i) / CDbl(rangeBlock.Rows.Count)
        
        ' ... and use that to calculate our new RGB ...
        red = CInt(toRed * ratio + fromRed * (1# - ratio))
        green = CInt(toGreen * ratio + fromGreen * (1# - ratio))
        blue = CInt(toBlue * ratio + fromBlue * (1# - ratio))
        
        ' ... and set the row, from the first to last column, in the new shade
        Range(Cells(rangeBlock.Row + i - 1, rangeBlock.Column), Cells(rangeBlock.Row + i - 1, rangeBlock.Columns.Count)).Interior.Color = RGB(red, green, blue)
    Next
End Sub

