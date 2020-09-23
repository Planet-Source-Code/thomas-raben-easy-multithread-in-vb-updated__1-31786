<div align="center">

## Easy multithread in VB \(Updated\)


</div>

### Description

Its actualy quite easy to make several threads in VB. All you need is one (1), API call and you're on your way.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-02-15 10:35:32
**By**             |[Thomas Raben](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-raben.md)
**Level**          |Intermediate
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Easy\_multi553822152002\.zip](https://github.com/Planet-Source-Code/thomas-raben-easy-multithread-in-vb-updated__1-31786/archive/master.zip)





### Source Code

To make serveral threads in VB, you need a single API call and a few constants (this code should be placed in a module)
:<br><br>
Public Const CTF_COINIT = &H8<br>
Public Const CTF_INSIST = &H1<br>
Public Const CTF_PROCESS_REF = &H4<br>
Public Const CTF_THREAD_REF = &H2<br><br>
Declare Function SHCreateThread Lib "shlwapi.dll" (ByVal pfnThreadProc As Long, pData As Any, ByVal dwFlags As Long, ByVal pfnCallback As Long) As Long<br><br>
Next is to make sub for the thread (it will not work for a private).<br><br>
public sub myNewThread()<br>
 dim i as integer<br><br>
 for i = 0 to 99<br>
 debug.print "message from a new thread (" & i & ")"<br>
 next i<br>
end sub<br><br>
Last but not least, all you need to do, is invoke the new thread:
SHCreateThread AddressOf myNewThread, ByVal 0&, CTF_INSIST, ByVal 0&<br><br>
That's it.<br><br>You have a new thread (make sure to end the sub before you exit the program or vb will crash.)

