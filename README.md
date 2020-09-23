<div align="center">

## 3d effects for any object\.


</div>

### Description

These codes produce 3d effects

on any form or picturebox.  The Etched3D and Raised3D

subs produce either a raised line or an etched line

around the picture box.  The SunckePanel3D and the

RaisedPanel3D subs produce a raised or lowered effect

on the entire form or picturebox. These look great in your bas.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[special](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/special.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/special-3d-effects-for-any-object__1-2573/archive/master.zip)





### Source Code

```
Public Sub SunkenPanel3D(obj As Object)
  ' Gives the effect of sinking the entire
  ' form or picture box, much like a 3d picture
  ' box with border style set to 1 - Fixed Single
  ' Hold the original scale mode
  Dim nScaleMode       As Integer
  ' Used for user defined scale only
  Dim sngScaleTop       As Single
  Dim sngScaleLeft      As Single
  Dim sngScaleWidth      As Single
  Dim sngScaleHeight     As Single
  If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
   nScaleMode = obj.ScaleMode
   If nScaleMode = 0 Then ' user defined scale
     sngScaleTop = obj.ScaleTop
     sngScaleLeft = obj.ScaleLeft
     sngScaleWidth = obj.ScaleWidth
     sngScaleHeight = obj.ScaleHeight
   End If
   obj.ScaleMode = 3 ' Pixel
   obj.Line (2, 2)-(obj.ScaleWidth - 1, 2), vb3DDKShadow
   obj.Line (2, 2)-(2, obj.ScaleHeight - 1), vb3DDKShadow
   obj.Line (2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, obj.ScaleHeight - 2), vb3DHighlight
   obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DHighlight
   ' Set the scale mode back to the same as it was
   obj.ScaleMode = nScaleMode
   If nScaleMode = 0 Then
     obj.ScaleTop = sngScaleTop
     obj.ScaleWidth = sngScaleWidth
     obj.ScaleLeft = sngScaleLeft
     obj.ScaleHeight = sngScaleHeight
   End If
  End If
End Sub
Public Sub RaisedPanel3D(obj As Object)
  ' Gives the effect of raising the entire
  ' picture box. Much like a 3d Panel
  ' Hold the original scale mode
  Dim nScaleMode       As Integer
  ' Used for user defined scale only
  Dim sngScaleTop       As Single
  Dim sngScaleLeft      As Single
  Dim sngScaleWidth      As Single
  Dim sngScaleHeight     As Single
  If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
   nScaleMode = obj.ScaleMode
   If nScaleMode = 0 Then ' user defined scale
     sngScaleTop = obj.ScaleTop
     sngScaleLeft = obj.ScaleLeft
     sngScaleWidth = obj.ScaleWidth
     sngScaleHeight = obj.ScaleHeight
   End If
   obj.ScaleMode = 3 ' Pixel
   obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DHighlight
   obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DHighlight
   obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DShadow
   obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DShadow
   ' Set the scale mode back to the same as it was
   obj.ScaleMode = nScaleMode
   If nScaleMode = 0 Then
     obj.ScaleTop = sngScaleTop
     obj.ScaleWidth = sngScaleWidth
     obj.ScaleLeft = sngScaleLeft
     obj.ScaleHeight = sngScaleHeight
   End If
  End If
End Sub
Public Sub Raised3D(obj As Object)
  ' Gives the effect of a raised line around
  ' the form or picturebox
  ' Hold the original scale mode
  Dim nScaleMode       As Integer
  ' Used for user defined scale only
  Dim sngScaleTop       As Single
  Dim sngScaleLeft      As Single
  Dim sngScaleWidth      As Single
  Dim sngScaleHeight     As Single
  If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
   nScaleMode = obj.ScaleMode
   If nScaleMode = 0 Then ' user defined scale
     sngScaleTop = obj.ScaleTop
     sngScaleLeft = obj.ScaleLeft
     sngScaleWidth = obj.ScaleWidth
     sngScaleHeight = obj.ScaleHeight
   End If
   obj.ScaleMode = 3 ' Pixel
   obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DHighlight
   obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DShadow
   obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DHighlight
   obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DShadow
   obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DHighlight
   obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DShadow
   obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DHighlight
   obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DShadow
   ' Set the scale mode back to the same as it was
   obj.ScaleMode = nScaleMode
   If nScaleMode = 0 Then
     obj.ScaleTop = sngScaleTop
     obj.ScaleWidth = sngScaleWidth
     obj.ScaleLeft = sngScaleLeft
     obj.ScaleHeight = sngScaleHeight
   End If
  End If
End Sub
Public Sub Etched3D(obj As Object)
  ' Gives the effect of an eteched line around the
  ' form or picture box.
  ' Hold the original scale mode
  Dim nScaleMode       As Integer
  ' Used for user defined scale only
  Dim sngScaleTop       As Single
  Dim sngScaleLeft      As Single
  Dim sngScaleWidth      As Single
  Dim sngScaleHeight     As Single
  If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
   nScaleMode = obj.ScaleMode
   If nScaleMode = 0 Then ' user defined scale
     sngScaleTop = obj.ScaleTop
     sngScaleLeft = obj.ScaleLeft
     sngScaleWidth = obj.ScaleWidth
     sngScaleHeight = obj.ScaleHeight
   End If
   obj.ScaleMode = 3 ' Pixel
   obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DShadow
   obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DHighlight
   obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DShadow
   obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DHighlight
   obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DShadow
   obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DHighlight
   obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DShadow
   obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DHighlight
   ' Set the scale mode back to the same as it was
   obj.ScaleMode = nScaleMode
   If nScaleMode = 0 Then
     obj.ScaleTop = sngScaleTop
     obj.ScaleWidth = sngScaleWidth
     obj.ScaleLeft = sngScaleLeft
     obj.ScaleHeight = sngScaleHeight
   End If
  End If
End Sub
```

