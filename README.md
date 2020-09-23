<div align="center">

## Fading backcolor


</div>

### Description

Make the backcolor property of a form acts like a VB-Setup Program
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Henrik Jacob Udsen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/henrik-jacob-udsen.md)
**Level**          |Unknown
**User Rating**    |4.3 (77 globes from 18 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/henrik-jacob-udsen-fading-backcolor__1-4605/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub Form_Paint()
 Dim lngY As Long
 Dim lngScaleHeight As Long
 Dim lngScaleWidth As Long
 Dim WhatColor As String
 ScaleMode = vbPixels
 lngScaleHeight = ScaleHeight
 lngScaleWidth = ScaleWidth
 DrawStyle = vbInvisible
 FillStyle = vbFSSolid
 For lngY = 0 To lngScaleHeight
  FillColor = RGB(0, 0, 255 - (lngY * 255) \ lngScaleHeight)
  Line (-1, lngY - 1)-(lngScaleWidth, lngY + 1), , B
 Next lngY
End Sub
```

