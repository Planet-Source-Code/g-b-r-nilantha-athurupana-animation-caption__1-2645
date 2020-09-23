<div align="center">

## Animation Caption\.


</div>

### Description

Quick and Easy! without API. Only a few lines of code.

You can scroll your Application's titles and form's caption.

I hope you enjoy.

G.B.R.Nilantha Athurupana.

Wattarama,Imbulgasdeniya,Sri Lanka.(94-03744479, E-Mail: crysbro@cga.lk)
 
### More Info
 
Create a new form (Any Name you can use). Add a Timer control (Timer1)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[G\.B\.R\.Nilantha Athurupana](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/g-b-r-nilantha-athurupana.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/g-b-r-nilantha-athurupana-animation-caption__1-2645/archive/master.zip)





### Source Code

```
'General Deciarations
Dim C As String 'to store current form's caption
Dim CO As Integer 'to store caption length
Dim FS As Long 'to store current form Width
Private Sub Form_Load()
 Timer1.Interval = 100
 Me.Caption = "Nilantha Athurupana"
 C = Me.Caption
 CO = Len(C) + 1
 Me.Caption = ""
 If Me.BorderStyle <> 2 Then
  FS = Me.ScaleWidth + 250
 Else
  FS = Me.ScaleWidth + 500
 End If
End Sub
Private Sub Form_Resize()
 If Me.WindowState = 1 Then
  FS = 3500
 Else
  FS = Me.ScaleWidth
 End If
End Sub
Private Sub Timer1_Timer()
On Error GoTo ATH
 Static C01 As Integer ' Counter 1
 Static CO2 As Integer ' Counter 2
 Static A As String 'To move Caption
 Dim R As String 'Restore Caption
 Dim T As String 'Restore Caption
XX:
 If CO > 0 Then
  C01 = CO
  T = Mid(C, C01, 1)
  CO = CO - 1
  R = " "
  Mid(R, 1) = T
  Me.Caption = R & Me.Caption
 Else
  A = A & " "
  R = " "
  Mid(R, 1) = A
  Me.Caption = R & Me.Caption
 End If
 If CO2 >= FS Then
  CO2 = 0
  CO = Len(C)
  Me.Caption = ""
  GoTo XX
 Else
  CO2 = CO2 + 50
 End If
 Exit Sub
ATH:
End Sub
```

