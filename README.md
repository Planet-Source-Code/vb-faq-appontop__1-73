<div align="center">

## AppOnTop


</div>

### Description

How do I get my application on top?To make your window truly topmost, use the SetWindowPos API call
 
### More Info
 
You can, to make the application stay on top, put the ZOrder method in a Timer event repeatedly called, say, every 1000 milliseconds. This makes a "softer" on-top than other methods, and allows the user to make a short peek below the form.

There are two different Zorder's of windows (forms) in Windows, both implemented internally as linked lists. One is for "normal" windows, the other for "topmost" windows (like the Clock application which is distributed with Windows). The ZOrder command above simply moves your window to the top of the "normal" window stack. There is another, independent stack for topmost windows - like those created by the example code above - which resolves problems if several of those should conflict.

Note that when a form is minimized, it loses its topmost attribute and you will have to set it again.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB FAQ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-faq.md)
**Level**          |Unknown
**User Rating**    |3.5 (21 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-faq-appontop__1-73/archive/master.zip)

### API Declarations

```

#IF WIN32 THEN
Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long
#ELSE 'Win16
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, _
ByVal hWndInsertAfter As Integer, ByVal X As Integer, _
ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, _
ByVal wFlags As Integer)
#END IF
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
```


### Source Code

```
To set Form1 as a top-most form, do the following:
#IF WIN32 THEN
Dim lResult as Long
lResult = SetWindowPos (me.hWnd, HWND_TOPMOST, _
0, 0, 0, 0, FLAGS)
#ELSE '16-bit API uses a Sub, not a Function
SetWindowPos me.hWnd, HWND_TOPMOST, _
0, 0, 0, 0, FLAGS
#END IF
To turn off topmost (make the form act normal again), do the following:
#IF WIN32 THEN
Dim lResult as Long
lResult = SetWindowPos (me.hWnd, HWND_NOTOPMOST, _
0, 0, 0, 0, FLAGS)
#ELSE '16-bit API uses a Sub, not a Function
SetWindowPos me.hWnd, HWND_NOTOPMOST, _
0, 0, 0, 0, FLAGS
#END IF
If you don't want to force a window on top, which will prevent the user from seeing below it, but simply want to move a Window to the top for the user's attention, do this:
Form1.ZOrder
```

