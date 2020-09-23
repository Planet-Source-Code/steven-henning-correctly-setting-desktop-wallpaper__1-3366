<div align="center">

## Correctly setting Desktop Wallpaper


</div>

### Description

I got tired of trying every wallpaper example on this site and finding they didn't work

on Windows NT.

This shows the example from Microsoft of how to correctly set the Desktop Wallpaper

from Visual Basic.

Original code is from :-

http://msdn.microsoft.com/library/techart/msdn_msdn192.htm

Seeing we are not allowed to link to another site, I just copied the code from

the above URL and modified it slightly.
 
### More Info
 
Add a Command Button control to Form1.

Command1 is created by default.

Set its Caption property to "Remove Wallpaper".

Add a second Command Button control to Form1.

Command2 is created by default.

Set its Caption property to "Change Wallpaper".


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steven Henning](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steven-henning.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steven-henning-correctly-setting-desktop-wallpaper__1-3366/archive/master.zip)

### API Declarations

```
Private Declare Function SystemParametersInfo Lib "user32" Alias _
  "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam _
  As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2
```


### Source Code

```
Private Sub Command1_Click()
  Dim X As Long
  X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, "(None)", _
	SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
  MsgBox "Wallpaper was removed"
End Sub
Private Sub Command2_Click()
  Dim FileName As String
  Dim X As Long
  ' Windows NT
  FileName = "c:\winnt\Coffee Bean.bmp"
  ' Windows 95 users, uncomment this line, you can delete the previous line
'  FileName = "c:\windows\Coffee Bean.bmp"
  X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, FileName, _
	SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
  MsgBox "Wallpaper was changed"
End Sub
```

