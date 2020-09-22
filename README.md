<div align="center">

## How to start a good page with ASP


</div>

### Description

Good information to help building a good ASP page.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bruno Martins Braga](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bruno-martins-braga.md)
**Level**          |Beginner
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bruno-martins-braga-how-to-start-a-good-page-with-asp__4-8580/archive/master.zip)





### Source Code

<p>@</p>
<p><font face="Verdana" size="2">First of all, we have to organize our page to easily find any part of the code we insered. It`s very important for you to put signs or usefull information in case you don`t remember what it was.</font></p>
<p align="center"><font color="#0000FF" size="2" face="Verdana">P.S.: This is a
tutorial, so please read carefully all the information before using any code.</font></p>
<p><font face="Verdana" size="2"><b>1) First step of the page</b><br>
</font></p>
<p><font color="#FF0000" size="2" face="Verdana">&lt;%@Language = "VBScript"%><br>
&lt;%<br>
Option Explicit<br>
response.expires = 0<br>
response.expiresabsolute = Now() -1<br>
response.addHeader "pragma","no-cache"<br>
response.addHeader "cache-control","private"<br>
Response.CacheControl = "no-cache"<br>
%></font></p>
<p><font face="Verdana" size="2">I always use a code in ASP, which has a particular function:</font></p>
<p><font face="Verdana" size="2">- <font color="#800000"> To force the browser to read the web page from the server, and not from the cache of the
computer</font>.&nbsp;<br>
For persons who are not pretty familiar with this, don`t worry, because s quite simple: everytime a page is loaded by a browser, depending on the configurations of couse, it will save some information about this site, specially if you see it with
frequently. If that happens, the next time you try to see the page, it will be
displayed the stored file from that address, not the real page from the server.
As ASP pages are known because they are dinamical, we would lose this
functionallity. This function just asks the browser to not keep any
&quot;cache&quot; inside the computer. Later, you will see that this same code
will help you to keep the &quot;Logout&quot; (particular page to make the log
out inside a web page) more secured.</font></p>
<p><font face="Verdana" size="2">- <font color="#800000">To request that all
variables to be Dimensioned</font>.<br>
It`s pretty usefull because it helps us to check if there is any variable we are
using by mistake, or incorrectly. t`s not a necessary command, but I advise, to
keep the programming sequence more organized.</font></p>
<p><font face="Verdana" size="2">Bu</font><font size="2">‚”</font><font face="Verdana" size="2">,
it also has its problems:</font></p>
<p><font face="Verdana" size="2">- I said that it will refresh everytime from
the server, not keeping any info inside the cache. The problem is that
everytime, the page will load all images again, even if they are the same. For
non fast internet connections, it might be a difference. But, anyway, if you
want something very dinamic, I suggest to use it, or the visitor won`t be able
to see your dinamic page...</font></p>
<p>@</p>
<p><b><font face="Verdana" size="2">2) Inserting the explanation</font></b></p>
<p><font face="Verdana" size="2">In ASP code, the symbol used to ignore a code
line is &quot; ' &quot;. But you can use it in any place, being aware of the
details:</font></p>
<p><font face="Verdana" size="2">- After you used this char, on the same line,
any code won`t be seen as a code, but as text.<br>
Eg:</font></p>
<p><font color="#FF0000" size="2" face="Verdana">&lt;%<br>
'This line is text<br>
<br>
Dim var1&nbsp;&nbsp;&nbsp;&nbsp; 'This var is for counting<br>
Dim var2&nbsp;&nbsp;&nbsp;&nbsp; 'This var is for catching<br>
%></font></p>
<p>@</p>
<p><b><font face="Verdana" size="2">3) Organizing the page functions - DIM</font></b></p>
<p><font face="Verdana" size="2">I prefer to guide all variables to the top of
the page, even before the HTML tags. It helps you not using the same variable
twice. In long programs, it may happen.</font></p>
<p><font face="Verdana" size="2">Another thing is to dimension a variable with a
characteristic name that will help you to catch its meaning right away. Avoid
using simple letters like &quot;i&quot; or &quot;var&quot;. It will become a
mess for long programming. Don`t be afraid of calling a variable as &quot;<font color="#FF0000">TextFromFormToBeChanged</font>&quot;.
The computer will accept it and you will never forget the idea of its value. As
you see, we usually try to put the first letter of a word in CAPS, to be easier
to read it later.</font></p>
<p>@</p>
<p><font face="Verdana" size="2"><b>4) Functions, Sequences &amp; Misc.</b></font></p>
<p><font face="Verdana" size="2">Just a part to say how much interest it is when
you organize your code in a way you can see the structure and recognise quickly
what is going on. For example, let`s see the command For - Next.</font></p>
<p><font color="#FF0000" size="2" face="Verdana">&lt;%<br>
For i=1 to 10<br>
response.write i<br>
Next<br>
%></font></p>
<p><font face="Verdana" size="2">When you see this code, it`s not quite
difficult to get its idea. It will print on screen the sequence of numbers from
1 to 10. And, if this code is inside some another code, you won`t mistake its
function. But, this is a case we usually don`t use, you should know this by now.
That`s why we try to separate dependent lines with the &quot;TAB&quot; button...
This would turn to:</font></p>
<p><font color="#FF0000" size="2" face="Verdana">&lt;%<br>
For i=1 to 10<br>
&nbsp;&nbsp;&nbsp; response.write i<br>
Next<br>
%></font></p>
<p><font face="Verdana" size="2">It does not seem to be changed a lot, but, with
long programming, believe, it`s very important.&nbsp;</font></p>
<p><font face="Verdana" size="2">I took one example I usually do in my websites,
to check this idea...</font></p>
<p><font face="Verdana" size="2">This next code check if a variable has a value,
then randomizes for a value that must be different from the first one. But don`t
worry with its funcionallity, that`s not that intention right now.</font></p>
<p><font color="#FF0000" size="2" face="Verdana">&lt;%<br>
If request("CorDefinida") = "" then&nbsp;<br>
	Randomize&nbsp;<br>
	CorRandom = Int((TotalCores * Rnd()) + 1)<br>
	if CorRandom = Session("CorSite") then<br>
		While CorRandom = session("CorSite")<br>
		randomize&nbsp;<br>
		CorRandom = Int((TotalCores * Rnd()) + 1)<br>
		Wend<br>
	end if<br>
	Session("CorSite") = Cores(CorRandom)<br>
else<br>
	Session("CorSite") = request("CorDefinida")<br>
end if<br>
%></font></p>
<p><font face="Verdana" size="2">But, look to the difference of reading this
code, and the next one...</font></p>
<p><font color="#FF0000" size="2" face="Verdana">&lt;%<br>
If request("CorDefinida") = "" then&nbsp;<br>
&nbsp;&nbsp;&nbsp; Randomize&nbsp;<br>
&nbsp;&nbsp;&nbsp; CorRandom = Int((TotalCores * Rnd()) + 1)<br>
&nbsp;&nbsp;&nbsp; if CorRandom = Session("CorSite") then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; While CorRandom = session("CorSite")<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Randomize&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; CorRandom = Int((TotalCores * Rnd()) + 1)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Wend<br>
&nbsp;&nbsp;&nbsp; end if<br>
&nbsp;&nbsp;&nbsp; Session("CorSite") = Cores(CorRandom)<br>
else<br>
&nbsp;&nbsp;&nbsp; Session("CorSite") = request("CorDefinida")<br>
end if<br>
%></font></p>
<p><font face="Verdana" size="2">What do you think? Makeing the code clear like
this, you help yourself getting bugs, and even improving the code later. </font></p>
<p><font face="Verdana" size="2">Happy programming...</font></p>

