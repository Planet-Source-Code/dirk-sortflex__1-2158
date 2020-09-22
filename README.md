<div align="center">

## SortFlex


</div>

### Description

Handle the sorting of a MSFlexgrid by only one sub-routine. Automatic ascenting and decending displayed by + and - in the Headline.
 
### More Info
 
syntax:

SortFlex MSFlexGrid, CollumToSort , StringSortAsBoolean , StringSortAsBoolean ...

example:

SortFlex flxProject, flxProject.MouseCol, False, True, True, True


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dirk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dirk.md)
**Level**          |Unknown
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dirk-sortflex__1-2158/archive/master.zip)





### Source Code

```
'
' 1999 by Dirk Bujna - b_dirk@yahoo.com
'
Public Sub SortFlex(FlexGrid As MSFlexGrid, TheCol As Integer, ParamArray IsString() As Variant)
  FlexGrid.Col = TheCol
  For i = 0 To FlexGrid.Cols - 1
    Headline = FlexGrid.TextMatrix(0, i)
    Ascend = Right$(Headline, 1) = "+"
    Decend = Right$(Headline, 1) = "-"
    If Ascend Or Decend Then Headline = Left$(Headline, Len(Headline) - 1)
    If i = TheCol Then
      If Ascend Then
        FlexGrid.TextMatrix(0, i) = Headline & "-"
        If IsMissing(IsString(i)) Then
          FlexGrid.Sort = flexSortGenericDescending
        Else
          If IsString(i) Then
            FlexGrid.Sort = flexSortStringDescending
          Else
            FlexGrid.Sort = flexSortNumericDescending
          End If
        End If
      Else
        FlexGrid.TextMatrix(0, i) = Headline & "+"
        If IsMissing(IsString(i)) Then
          FlexGrid.Sort = flexSortGenericAscending
        Else
          If IsString(i) Then
            FlexGrid.Sort = flexSortStringAscending
          Else
            FlexGrid.Sort = flexSortNumericAscending
          End If
        End If
      End If
    Else
      FlexGrid.TextMatrix(0, i) = Headline
    End If
  Next i
End Sub
```

