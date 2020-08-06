<%
Function UBBCode(strContent)
		strContent=Replace(strContent,"[#seperator#]","")
		Dim re, strMatches, strMatch, tmpStr1, tmpStr2, tmpStr3, tmpStr4, RNDStr
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
   		re.Pattern="(javascript)"
		strContent=re.Replace(strContent,"&#106avascript")
		re.Pattern="(jscript:)"
		strContent=re.Replace(strContent,"&#106script:")
		re.Pattern="(js:)"
		strContent=re.Replace(strContent,"&#106s:")
		re.Pattern="(value)"
		strContent=re.Replace(strContent,"&#118alue")
		re.Pattern="(about:)"
		strContent=re.Replace(strContent,"about&#58")
		re.Pattern="(file:)"
		strContent=re.Replace(strContent,"file&#58")
		re.Pattern="(document.cookie)"
		strContent=re.Replace(strContent,"documents&#46cookie")
		re.Pattern="(vbscript:)"
		strContent=re.Replace(strContent,"&#118bscript:")
		re.Pattern="(vbs:)"
		strContent=re.Replace(strContent,"&#118bs:")
		re.Pattern="(on(mouse|exit|error|click|key))"
		strContent=re.Replace(strContent,"&#111n$2")

		'自动识别www等开头的网址
		re.Pattern="([^=\]][\s]*?|^)(https?|ftp|gopher|news|telnet|mms|rtsp)://([a-z0-9/\-_+=.~!%@?#%&;:$\\()|]+)"
		StrContent=re.Replace(StrContent,"$1<a href=$2://$3>$2://$3</a>")

		'If Enable_Emot = 1 Then   '表情开关
			re.Pattern="(\[emot([0-9]{2})\])"
			strContent=re.Replace(strContent,"<img src=""images/emot/$2.gif"" class=""emot"" />")
		'End If
		re.Pattern="\[code\](<br>)+"
		strContent=re.Replace(strContent,"[code]")
		re.Pattern="\[quote\](<br>)+"
		strContent=re.Replace(strContent,"[quote]")
		re.Pattern="\[img\](.*?)\[\/img\]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
			strContent=Replace(strContent,strMatch.Value,"<img onerror=""javascript:errpic(this)"" onload=""javascript:if(this.width>500)(this.width=500);"" src="""&tmpStr1&""" border=""0""   alt=""&#25353;&#27492;&#22312;&#26032;&#31383;&#21475;&#25171;&#24320;&#22270;&#29255;"" onmouseover=""this.style.cursor='hand';"" onmousewheel=""return bbimg(this)"" onclick=""window.open(this.src);"">")
		Next
		Set strMatches=Nothing
		re.Pattern="\[img=(left|right|center|absmiddle)\](.*?)\[\/img\]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
		tmpStr1=strMatch.SubMatches(0)
		tmpStr2=CheckLinkStr(strMatch.SubMatches(1))
		strContent=Replace(strContent,strMatch.Value,"<img class='ubbimg' src="""&tmpStr2&""" align="""&tmpStr1&"""  border=""0""  alt=""&#25353;&#27492;&#22312;&#26032;&#31383;&#21475;&#25171;&#24320;&#22270;&#29255;"" onmouseover=""this.style.cursor='hand';"" onclick=""window.open(this.src);"">")
		Next
		Set strMatches=Nothing
		If Enable_UBB_Media = 1 Then    '媒体开关
		strContent=replace(strContent,"[swf]","[swf=720,524]")
		strContent=replace(strContent,"[wmv]","[wmv=550,400]")
		strContent=replace(strContent,"[wma]","[wma=450,70]")
		strContent=replace(strContent,"[rm]","[rm=550,400]")
		strContent=replace(strContent,"[ra]","[ra=450,60]")
		re.Pattern="\[(swf|wma|wmv|rm|ra|qt)=(\d*?|),(\d*?|)\](.*?)\[\/(swf|wma|wmv|rm|ra|qt)\]"
'		re.Pattern="\[(wma|wmv|rm|ra|qt)=(\d*?|),(\d*?|)\](.*?)\[\/(wma|wmv|rm|ra|qt)\]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch in strMatches
			RNDStr=Int(7999 * Rnd + 2000)
			tmpStr1=CheckLinkStr(strMatch.SubMatches(3))
			strContent= Replace(strContent,strMatch.Value,"<div class=""code_head""><input id=""VOBJ_"&RNDStr&""" type=""hidden"" value=""-1"" /><a href=""javascript:UBBShowObj('"&strMatch.SubMatches(0)&"','OBJ_"&RNDStr&"','"&strMatch.SubMatches(3)&"','"&strMatch.SubMatches(1)&"','"&strMatch.SubMatches(2)&"');""><img src=""images/icon_media.gif"" alt=""&#26174;&#31034;&#24433;&#38899;&#25991;&#20214;"" align=""absmiddle"" border=""0"" />打开/关闭媒体</a></div><br><div id=""OBJ_"&RNDStr&"""></div>")
		Next
		Set strMatches=Nothing
		End If
		re.Pattern = "\[url=(.[^\]]*)\](.*?)\[\/url]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
			tmpStr2=strMatch.SubMatches(1)
			strContent=Replace(strContent,strMatch.Value,"<img align=absmiddle src=images/url.gif><a target=""_blank"" href="""&tmpStr1&"""><u>"&tmpStr2&"</u></a>")
		Next
		Set strMatches=Nothing
	
		re.Pattern = "\[url](.*?)\[\/url]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			tmpStr1=CheckLinkStr(strMatch.SubMatches(0))
			tmpStr2=CutURL(tmpStr1)
			strContent=Replace(strContent,strMatch.Value,"<a target=""_blank"" href="""&tmpStr1&"""><u>"&tmpStr2&"</u></a>")
		Next
		Set strMatches=Nothing

		re.Pattern = "\[email=(.[^\]]*)\](.*?)\[\/email]"
		strContent = re.Replace(strContent,"<img align=absmiddle src=images/email.gif><a href=""mailto:$1"">$2</a>")
		re.Pattern = "\[email](.*?)\[\/email]"
		strContent = re.Replace(strContent,"<img align=absmiddle src=images/email.gif><a href=""mailto:$1"">$1</a>")
		re.Pattern="\[fly\](.*)\[\/fly\]"
		strContent=re.Replace(strContent,"<marquee width=100% behavior=alternate scrollamount=3>$1</marquee>")
		re.Pattern="\[move\](.*)\[\/move\]"
		strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$1</marquee>")	
		re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
		strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:shadow(color=$2,strength=$3)"">$4</td></tr></table>")
		re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
		strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:glow(color=$2,strength=$3)"">$4</td></tr></table>")
         strContent = Replace(strContent,"[sub]","<sub>")
		strContent = Replace(strContent,"[/sub]","</sub>")
		strContent = Replace(strContent,"[sup]","<sup>")
		strContent = Replace(strContent,"[/sup]","</sup>")

		strContent = Replace(strContent,"[list]","<ul>")
		strContent = Replace(strContent,"[list=1]","<ol type=""1"">")
		strContent = Replace(strContent,"[list=a]","<ol type=""a"">")
		strContent = Replace(strContent,"[list=A]","<ol type=""A"">")
		strContent = Replace(strContent,"[*]","<li>")
		strContent = Replace(strContent,"[/list]","</ul></ol>")
		
		re.Pattern="\[font=([^<>\]]*?)\](.*?)\[\/font]"
		strContent=re.Replace(strContent,"<font face=""$1"">$2</font>")
		re.Pattern="\[color=([^<>\]]*?)\](.*?)\[\/color]"
		strContent=re.Replace(strContent,"<font color=""$1"">$2</font>")
		re.Pattern="\[align=([^<>\]]*?)\](.*?)\[\/align]"
		strContent=re.Replace(strContent,"<div align=""$1"">$2</div>")
		re.Pattern="\[size=(\d*?)\](.*?)\[\/size]"
		strContent=re.Replace(strContent,"<font size=""$1"">$2</font>")
		re.Pattern="\[b\](.*?)\[\/b]"
		strContent=re.Replace(strContent,"<strong>$1</strong>")	
		re.Pattern="\[i\](.*?)\[\/i]"
		strContent=re.Replace(strContent,"<em>$1</em>")	
		re.Pattern="\[u\](.*?)\[\/u]"
		strContent=re.Replace(strContent,"<u>$1</u>")

		re.Pattern="\[code\](.*?)\[\/code\]"
		Set strMatches=re.Execute(strContent)
		For Each strMatch In strMatches
			RNDStr=Int(7999 * Rnd + 2000)
			tmpStr1=strMatch.SubMatches(0)
			strContent= Replace(strContent,strMatch.Value,"<div class=""code_head"" style=""float:left"">程序代码:</div><div class=""code_head"" style=""float:right""><a href=""javascript:CopyText(document.all.CODE_"&RNDStr&");"">[ 复制代码 ]</a> <a href=""javascript:runCode(document.all.CODE_"&RNDStr&");"">[ 运行代码 ]</a></div><br><div class=""code_main"" id=""CODE_"&RNDStr&""" style=""overflow-y:auto;overflow-x:hidden;height:200px;"">"&tmpStr1&"</div>")
		Next
		Set strMatches=Nothing

		re.Pattern="\[quote\](.*?)\[\/quote\]"
		strContent= re.Replace(strContent,"<div class=""code_head"">&#24341;&#29992;&#20869;&#23481;&#65306;</div><div class=""code_main"">$1</div>")
		
		If User_ID<>Empty Then
			re.Pattern="\[mem\](.*?)\[\/mem\]"
			strContent= re.Replace(strContent,"<div class=""code_main""><font color=#9D9D9D>以下是注册用户才能看到的内容</font><BR><BR>$1<BR><BR></div>")
		ELSE
			re.Pattern="\[mem\](.*?)\[\/mem\]"
			strContent= re.Replace(strContent,"<BR><div class=""code_main""> <b>以下内容只有<a href=""register.asp""><font color=#ff0000>注册用户</font></a>才能看到</b></div><BR>")
		END IF

		re.Pattern = "\[down=(.[^\]]*)\](.*?)\[\/down]"
		strContent = re.Replace(strContent,"<img src=""images/download.gif"" align=""absmiddle"" /> <a href=""$1"" target=""_blank"">$2</a>")

		re.Pattern="\[quote2\](.*?)\[\/quote2\]"
		strContent= re.Replace(strContent,"<div class=""code_main"">$1</div>")


		Set re=Nothing
        'strContent=replace(replace(strContent,"[page]",""),"[/page]","")
		UBBCode=strContent

End Function

Function CheckLinkStr(Str)
	Str = Replace(Str, "document.cookie", ".")
	Str = Replace(Str, "document.write", ".")
	Str = Replace(Str, "javascript:", "javascript ")
	Str = Replace(Str, "vbscript:", "vbscript ")
	Str = Replace(Str, "javascript :", "javascript ")
	Str = Replace(Str, "vbscript :", "vbscript ")
	Str = Replace(Str, "[", "&#91;")
	Str = Replace(Str, "]", "&#93;")
	Str = Replace(Str, "<", "&#60;")
	Str = Replace(Str, ">", "&#62;")
	Str = Replace(Str, "{", "&#123;")
	Str = Replace(Str, "}", "&#125;")
	Str = Replace(Str, "|", "&#124;")
	Str = Replace(Str, "script", "&#115;cript")
	Str = Replace(Str, "SCRIPT", "&#083;CRIPT")
	Str = Replace(Str, "Script", "&#083;cript")
	Str = Replace(Str, "script", "&#083;cript")
	Str = Replace(Str, "object", "&#111;bject")
	Str = Replace(Str, "OBJECT", "&#079;BJECT")
	Str = Replace(Str, "Object", "&#079;bject")
	Str = Replace(Str, "object", "&#079;bject")
	Str = Replace(Str, "applet", "&#097;pplet")
	Str = Replace(Str, "APPLET", "&#065;PPLET")
	Str = Replace(Str, "Applet", "&#065;pplet")
	Str = Replace(Str, "applet", "&#065;pplet")
	Str = Replace(Str, "embed", "&#101;mbed")
	Str = Replace(Str, "EMBED", "&#069;MBED")
	Str = Replace(Str, "Embed", "&#069;mbed")
	Str = Replace(Str, "embed", "&#069;mbed")
	Str = Replace(Str, "document", "&#100;ocument")
	Str = Replace(Str, "DOCUMENT", "&#068;OCUMENT")
	Str = Replace(Str, "Document", "&#068;ocument")
	Str = Replace(Str, "document", "&#068;ocument")
	Str = Replace(Str, "cookie", "&#099;ookie")
	Str = Replace(Str, "COOKIE", "&#067;OOKIE")
	Str = Replace(Str, "Cookie", "&#067;ookie")
	Str = Replace(Str, "cookie", "&#067;ookie")
	Str = Replace(Str, "event", "&#101;vent")
	Str = Replace(Str, "EVENT", "&#069;VENT")
	Str = Replace(Str, "Event", "&#069;vent")
	Str = Replace(Str, "event", "&#069;vent")
	CheckLinkStr = Str
End Function

function dvHTMLEncode(byval fString)
	if isnull(fString) or trim(fString)="" then
		dvHTMLEncode=""
		exit function
	end if
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    dvHTMLEncode = fString
end function

function dvHTMLCode(byval fString)
	if isnull(fString) or trim(fString)="" then
		dvHTMLCode=""
		exit function
	end if
    fString = replace(fString, "&gt;", ">")
    fString = replace(fString, "&lt;", "<")

    fString = Replace(fString,  "&nbsp;"," ")
    fString = Replace(fString, "&quot;", CHR(34))
    fString = Replace(fString, "&#39;", CHR(39))
    fString = Replace(fString, "</P><P> ",CHR(10) & CHR(10))
    fString = Replace(fString, "<BR> ", CHR(10))

    dvHTMLCode = fString
end function

Function CutURL(URLStr)
	Dim URLLen
	URLLen=Len(URLStr)
	If URLLen>65 Then
		CutURL=Left(URLStr,URLLen*0.5)&" ... "&Right(URLStr,URLLen*0.3)
	Else
		CutURL=URLStr
	End If
End Function

%>