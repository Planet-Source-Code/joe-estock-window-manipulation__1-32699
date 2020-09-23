<div align="center">

## Window Manipulation


</div>

### Description

This code will show you how to retreive a handle of an open window and how to manipulate that window once you have it's handle.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joe Estock](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joe-estock.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joe-estock-window-manipulation__1-32699/archive/master.zip)

### API Declarations

```
'Add the following declarations to your form's
'General (Declarations) section
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
```


### Source Code

```
'Add a command button to a form, set the command
'button's name to Command1, then add this code
'to your form's code.
Private Sub Command1_Click()
  Dim lhWnd As Long  'Holds the handle to the window
  'The FindWindow parameter is as follows:
  'lpClassName:  This is the name of the class
  '        Use the Spy ++ utility to retreive
  '        this information
  'lpWindowName: This is the caption of the window
  lhWnd = FindWindow("Minesweeper", "Minesweeper")
  Text1.Text = lhWnd
  'Make sure we have a valid handle
  If lhWnd <> 0 Then
    'Flash the window by inverting it
    'If it has focus, then remove focus
    'If it doesn't have focus, give it focus
    FlashWindow lhWnd, 1
  End If
  If lhWnd <> 0 Then
    'Change the window's caption
    SetWindowText lhWnd, "ChangedSweeper"
  End If
End Sub
```

