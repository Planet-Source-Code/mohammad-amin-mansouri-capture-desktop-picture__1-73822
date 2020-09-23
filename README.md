<div align="center">

## Capture Desktop Picture 


</div>

### Description

Capture Desktop Picture (*.jpg")
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mohammad Amin Mansouri](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mohammad-amin-mansouri.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mohammad-amin-mansouri-capture-desktop-picture__1-73822/archive/master.zip)





### Source Code

```
'//Capture Desktop
'Programmer : Mohammad Amin Mansouri :p
'forum : Wwww.Forum.Honarjo.com
'WebSite : Www.AhoraChat.Net & Www.Iridiver.Net
'Phone: +989390763223
'Email : SetareSokhte@Gmail.com
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim Img As ImageFile
Dim IP As ImageProcess
Dim Pic As ImageFile
Private Sub Form_Load()
Call CaptureDesktop("c:\amin.jpg")
End Sub
Public Function CaptureDesktop(Path As String) As String
Set Img = New ImageFile
Set IP = New ImageProcess
IP.Filters.Add IP.FilterInfos("Convert").FilterID
IP.Filters(1).Properties(1).Value = wiaFormatJPEG
Form1.AutoRedraw = True
Form1.ScaleMode = vbpixel
Amin = GetDesktopWindow()
Persian = GetDC(Amin)
BitBlt Form1.hDC, 0, 0, Form1.Width, Form1.Height, Persian, 0, 0, vbSrcCopy
SavePicture Form1.Image, ("c:\Persian.bmp")
Img.LoadFile ("c:\Persian.bmp")
Set Pic = IP.Apply(Img)
Pic.SaveFile (Path)
Kill ("c:\persian.bmp")
End Function
```

