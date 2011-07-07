<%

Dim TBoxVal
Dim SessVal

TBoxVal = Request.Form("CaptchaStr")
SessVal = Session("Captcha")

Response.write "Session Value: " & Sessval & "<br><br>"
Response.write "Form Value: " & TBoxVal  & "<br><br>"

If CStr(SessVal) = CStr(TBoxVal) then
	Response.write "<h2><font color=blue> Captcha Match !!! <p> U r an HUMAN !!!</font></h2>"
else
	Response.write "<h2><font color=red> Captcha MisMatch !!! U r a BOT !!!</font></h2>"
end if

%>