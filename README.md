<div align="center">

## Quick N\! \(Factorial\) calculation using Double\!


</div>

### Description

Explain a metodology for a quick computation of

N! ( factorial) suggestion and optimization are wellcomed!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[michele berardi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michele-berardi.md)
**Level**          |Advanced
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michele-berardi-quick-n-factorial-calculation-using-double__1-38817/archive/master.zip)





### Source Code

```
Dim N As Double, b As Double, c As Double, p As Double
'
' Fast Way To Calculate N! ( N Factorial)
'
' A is N
' using visual basic my original algorithm
' is adapted to the vb limits..
' you can use long or int instead of double
' for small calculation..
' some tips require the use of asr (aritmetic
' shift right)
' instead of division by 2!
' and code optimization instead - 1 you can....
' a good exercize of optimization... enjoy!
' (also assembly form of this code boost the
' performances!)
'
' N.B.
' using double I extend the range of N!
' that i can represent!
' PASS TO VARIABLE N
' THE VALUE FOR WITCH
' YOU MUST CALCULATE
' FACTORIAL ( N! )
'
c = N - 1
p = 1
While c > 0
p = 0
b = c
While b > 0
If b And 1 Then
p = p + N
End If
b = int (b / 2) ' YOU MUST USE THE INTEGER PART NOT THE REST! asr more efficient fo division!
N = N + N
Wend
N = p
c = c - 1
Wend
MsgBox p ' the result of N!
```

