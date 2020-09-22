<div align="center">

## Update \- Close Window \(partial title\)


</div>

### Description

This will get the hWnd of freaky programs that change class names every time you start it, and then close it. The only thing you need to know is part of the Titlename (Netsc, instead of "Web Page Title" - Netscape). This WILL close Microsoft Internet Explorer (old version wouldn't) now =Þ. Also, I made it so that you don't have to have a form named "Form1" (by using EnumWindows)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Che](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/che.md)
**Level**          |Beginner
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/che-update-close-window-partial-title__1-8784/archive/master.zip)

### API Declarations

```
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_CLOSE = &H10
Public AppTitle As String
Public ApphWnd As Long
```


### Source Code

```
Function GetCaption(WindowhWnd)
  hwndlength% = GetWindowTextLength(WindowhWnd)
  hWndTitle$ = String$(hwndlength%, 0)
  a% = GetWindowText(WindowhWnd, hWndTitle$, (hwndlength% + 1))
  GetCaption = hWndTitle$
End Function
Function CheckAllWindows(ByVal hwnd As Long, lParam As Long) As Boolean
  Dim a
  a = LCase(GetCaption(hwnd))
  If InStr(1, a, LCase(AppTitle)) <> 0 Then
    ApphWnd = hwnd
    CheckAllWindows = False
  Else
    CheckAllWindows = True
  End If
End Function
Sub KillWin(Title As String)
  Dim a
  AppTitle = Title
  EnumWindows AddressOf CheckAllWindows, 0&
  If ApphWnd = 0 Then Exit Sub
  a = PostMessage(ApphWnd, WM_CLOSE, 0&, 0&)
End Sub
'-----Use KillWin to close the window. KillWin "Title"
```

