<div align="center">

## Small tips: validate control array index


</div>

### Description

<p>I've seen many ways to check whether a certain control index exists in a control array, where all indexes are not filled. For example, you could have a control array of labels with indexes 0, 1 & 3. I guess one of the most common ways to check for this is to use <b>On Error</b>, but I've always found resorting to rising errors a bad habit to do. A much better another alternative is to loop through the control array using For Each and check for the Index property, but even this becomes a bit silly if you know the index you want to check for and loop the entire control array just to check if it is there or not.</p>

<p>The solution: <b>VarType()</b></p>

<p>VarType returns the variable type of default property of a control, ie. <b>VarType(MyLabels(1))</b> returns 8 (vbString), from Caption property. What makes this usable for sniffing for valid index is that VarType returns 9 (vbObject) for invalid indexes! There is no control that returns an object in their default property, making this a perfectly safe method to check if an index exists or not.</p>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vesa Piittinen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vesa-piittinen.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vesa-piittinen-small-tips-validate-control-array-index__1-73714/archive/master.zip)





### Source Code

```
' make sure you add MyLabel to your form and set Index property to 0
Option Explicit
Private Sub Form_Load()
  ' control array indexes are Integers
  Dim I As Integer
  Dim C As Label
  Load MyLabel(1)
  Load MyLabel(3)
  ' solution one: On Error... looping through all controls
  On Error Resume Next
  For I = MyLabel.LBound To MyLabel.UBound
    ' check if the index is there
    MyLabel(I) = MyLabel(I)
    If Err = 0 Then
      ' do what you need to do...
    Else
      Debug.Print "On Error, Invalid index: " & I
      Err.Clear
    End If
  Next I
  ' that wasn't pretty...
  On Error GoTo 0
  ' solution two: For Each for specific index, must loop through all controls
  I = 2
  For Each C In MyLabel
    ' if the index is there we exit the loop
    If C.Index = I Then Exit For
  Next C
  ' if we passed though all controls then C is now Nothing
  If Not C Is Nothing Then
    Debug.Print "For Each, VALID index: " & I
  Else
    Debug.Print "For Each, Invalid index: " & I
  End If
  ' solution three: VarType looping through all controls
  For I = MyLabel.LBound To MyLabel.UBound
    If VarType(MyLabel(I)) <> vbObject Then
      ' VALID
    Else
      Debug.Print "VarType, Invalid index: " & I
    End If
  Next I
  ' solution four: VarType for specific index
  I = 2
  If VarType(MyLabel(I)) = vbObject Then Debug.Print "VarType, Invalid index: " & I
End Sub
```

