<div align="center">

## Splitting Strings


</div>

### Description

The purpose of this article is to teach the reader how to split a string into fragments and use the fragments effectively. Useful for Winsock programs, date commands, and other possible needs.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Particle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/particle.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/particle-splitting-strings__1-32089/archive/master.zip)





### Source Code

<p>In order to split a string, you'll need a source. An internal string or a
textbox (and so on; any major text source) will do fine. In addition to the
source, you'll need something to use it. Look at the code below for a the raw
code:<br>
<br>
output = Split(source, split object)(part number)<br>
<br>
If you wanted to split a field called Input into two sections, then you'd do the
following:<br>
<br>
<font color="#003399">Private Sub Command1_Click()</font><br>
&nbsp;&nbsp;&nbsp; Dim InputString as String<br>
&nbsp;&nbsp;&nbsp; Dim Output1 as String<br>
&nbsp;&nbsp;&nbsp; Dim Output2 as String<br>
<br>
&nbsp;&nbsp;&nbsp; 'Sets the input field to a sample string that could be used
in a chat program for example<br>
&nbsp;&nbsp;&nbsp; InputString = &quot;Particle:::Hello bob!&quot;<br>
<br>
&nbsp;&nbsp;&nbsp; Output1 = Split(InputString, &quot;:::&quot;)(0)<br>
&nbsp;&nbsp;&nbsp; Output2 = Split(InputString, &quot;:::&quot;)(1)<br>
<font color="#003399">Exit Sub</font><br>
<br>
Alright, do you see how it works? You basically identify what you want split,
what sequence to split it at, and which segment to set that particular output
string to use. You could go on further with the segment number for if you had
something like this perhaps: &quot;Particle:::How's it goin?:::Yo&quot;.<br>
<br>
That code would yield the following:<br>
<br>
Output1 = &quot;Particle&quot;<br>
Output2 = &quot;Hello bob!&quot;<br>
<br>
Notice that the splitting sequence is removed. Add the following code before the
rest if your string may not always have the same number of segments:<br>
<br>
On Error Resume Next<br>
<br>
FINAL SAMPLE:<br>
<font color="#003399">Private Sub Command1_Click()</font><br>
&nbsp;&nbsp;&nbsp; Dim InputString as String<br>
&nbsp;&nbsp;&nbsp; Dim Output1 as String<br>
&nbsp;&nbsp;&nbsp; Dim Output2 as String<br>
&nbsp;&nbsp;&nbsp; Dim Output3 as String<br>
<br>
&nbsp;&nbsp;&nbsp; 'Sets the input field to a sample string that could be used
in a chat program for example<br>
&nbsp;&nbsp;&nbsp; InputString = &quot;Particle:::Hello bob!:::Yo&quot;<br>
<br>
&nbsp;&nbsp;&nbsp; Output1 = Split(InputString, &quot;:::&quot;)(0)<br>
&nbsp;&nbsp;&nbsp; Output2 = Split(InputString, &quot;:::&quot;)(1)<br>
&nbsp;&nbsp;&nbsp; Output3 = Split(InputString, &quot;:::&quot;)(2)<br>
<font color="#003399">Exit Sub</font><br>
<br>
You don't have to use every segment either. You could use segments 0, 4, and 11
for example.<br>
<br>
This is also very handy for splitting the Time command, identifying things for
winsock applications etc. Use it in the same way, splitting on the &quot;:&quot;s for
example.</p>

