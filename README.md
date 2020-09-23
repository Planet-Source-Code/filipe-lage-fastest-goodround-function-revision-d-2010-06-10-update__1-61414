<div align="center">

## FASTEST GoodRound function \(Revision D\) 2010\-06\-10 Update\!


</div>

### Description

Provides a good mathematical rounding of numbers instead of VB's "banking" round function.<br>

' Revision C by Donald - 20060201 - (Bugfix)<br>

' Revision D by Jeroen De Maeijer - 20100529 - (Bugfix)<br>

' Revision E by Filipe Lage - 20100530 (speed improvements)<br>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Filipe Lage](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/filipe-lage.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/filipe-lage-fastest-goodround-function-revision-d-2010-06-10-update__1-61414/archive/master.zip)





### Source Code

<br><br>
<p><font face=Verdana size=2><strong>As you probably know, VB6 round function doesn't provide a good mathematical rounding of numbers, nor does it support negative rounding decimals</strong><br>
Example: Round(2.5) results in 2 instead of the right value: 3<br>
Another example is that round doesn't support negative rounding decimals: Example: a Round of '1100' with decimalcases '-2' should result in 1000 and in VB internal round function, it fails with an error.<br>
VBSpeed site (http://www.xbeat.net/vbspeed/) provides many examples and benchmarks of solutions to this problem, so I've decided to submit 2 source codes providing the SHORTEST good round function and the FASTEST good round function.<br>
In this source code I'll provide the FASTEST (for 5 years and counting) to return the right numeric methematical rounding of a number, including support for negative rounding.<br>
<br/>
You can check benchmarks at http://www.xbeat.net/vbspeed/c_Round.htm (VBSpeed Round source codes)</a>.<br>
New revision D, with bugfix and even faster.<br>
So, here it is:
</font></p>
<br>
<font face=Arial size=1>
<br>
Public Function GoodRound(ByVal v As Double, Optional ByVal lngDecimals As Long = 0) As Double<br>
 ' By Filipe Lage<br>
 ' fclage@gmail.com<br>
 ' msn: fclage@clix.pt<br>
 ' Revision C by Donald - 20060201 - (Bugfix)<br>
 ' Revision D by Jeroen De Maeijer - 20100529 - (Bugfix)<br>
 ' Revision E by Filipe Lage - 20100530 (speed improvements)<br>
 Dim xint As Double, yint As Double, xrest As Double<br>
 Static PreviousValue  As Double<br>
 Static PreviousDecimals As Long<br>
 Static PreviousOutput  As Double<br>
 Static M        As Double<br>
   <br>
 If PreviousValue = v And PreviousDecimals = lngDecimals Then GoodRound = PreviousOutput: Exit Function<br>
   ' Hey... it's the same number and decimals as before...<br>
   ' So, the actual result is the same. No need to recalc it<br>
 <br>
 If v = 0 Then Exit Function<br>
   ' no matter what rounding is made, 0 is always rounded to 0<br>
   <br>
 If PreviousDecimals = lngDecimals Then<br>
   ' 20100530 Improvement by fclage - Moved M initialization here for speedup<br>
   If M = 0 Then M = 1 ' Initialization - M is never 0 (it is always 10 ^ n)<br>
   Else<br>
   ' A different number of decimal places, means a new Multiplier<br>
   PreviousDecimals = lngDecimals<br>
   M = 10 ^ lngDecimals<br>
   End If<br>
 <br>
 If M = 1 Then xint = v Else xint = v * CDec(M)<br>
   ' Let's consider the multiplication of the number by the multiplier<br>
   ' Bug fixed: If you just multiplied the value by M, those nasty reals came up<br>
   ' So, we use CDEC(m) to avoid that<br>
                               <br>
 GoodRound = Fix(xint)<br>
   ' The real integer of the number (unlike INT, FIX reports the actual number)<br>
 <br>
 ' 20060201: fix by Donald<br>
 If Abs(Fix(10 * (xint - GoodRound))) > 4 Then<br>
  If xint < 0 Then '20100529 fix by Zoenie:<br>
  ' previous code would round -0,0714285714 with 1 decimal in the end result to 0.1 !!!<br>
  ' 20100530 Speed improvement by Filipe - comparing vars with < instead of >=<br>
   GoodRound = GoodRound - 1<br>
  Else<br>
   GoodRound = GoodRound + 1<br>
  End If<br>
 End If<br>
   ' First decimal is 5 or bigger ? If so, we'll add +1 or -1 to the result (later to be divided by M)<br>
 <br>
 If M = 1 Then Else GoodRound = GoodRound / M<br>
   ' Divides by the multiplier. But we only need to divide if M isn't 1<br>
 <br>
 PreviousOutput = GoodRound<br>
 PreviousValue = v<br>
   ' Let's save this last result in memory... may be handy ;)<br>
End Function

