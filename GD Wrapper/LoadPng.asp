<%@language=jscript%><%
Response.ContentType = "image/png";
var gdImage = Server.CreateObject("GDLibrary.gdImage");
if (!gdImage.LoadFromFile(Server.MapPath("test.png"))) {
	gdImage.Create(200, 20);
	gdImage.ColorAllocate(0, 0, 0);
	gdImage.Chars(gdImage.FontGetMediumBold(), 25, 5, "Image Not Found", gdImage.ColorAllocate(255, 255, 255));
}
Response.BinaryWrite(gdImage.ToPngStream().Read);
%>