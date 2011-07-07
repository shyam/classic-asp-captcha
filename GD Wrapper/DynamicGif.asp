<%@language=jscript%><%

// Write image/gif header
Response.ContentType = "image/gif";

// Set 'dynamic' content to expire immediately (depends on situation)
Response.Expires = -1000;

// Create variables
var gdImage = Server.CreateObject("GDLibrary.gdImage");

// Create a new paleted image in memory
//gdImage.Create(250, 200);

// Because we're gonna draw a gradient, we create a True Color image!
// GD Library will dither and convert it to a Palette on output!
gdImage.CreateTrueColor(250, 200);

// Add/Return Colors to Palette
var Black = gdImage.ColorAllocate(0, 0, 0); // First color added will always be background color. Palette Index = 0
var Red = gdImage.ColorAllocate(255, 0, 0); // Palette Index = 1
var Green = gdImage.ColorAllocate(0, 255, 0); // Palette Index = 2
var Blue = gdImage.ColorAllocate(0, 0, 255); // Palette Index = 3 (Image is now 4 color)
var Yellow = gdImage.ColorAllocate(255, 255, 0); // Image becomes 8 color
var Magenta = gdImage.ColorAllocate(255, 0, 255);
//var RandomColor1 = gdImage.ColorAllocate(Math.random() * 255, Math.random() * 255, Math.random() * 255);
//var RandomColor2 = gdImage.ColorAllocate(Math.random() * 255, Math.random() * 255, Math.random() * 255);

// Drawing 3 Squares
gdImage.Rectangle(0, 0, 5, 5, Red);
gdImage.Rectangle(10, 10, 15, 15, Green);
gdImage.Rectangle(20, 20, 25, 25, Blue);

// Draw 300 Random Pixels
for (var i = 0; i < 100; i++) {
	gdImage.SetPixel(Math.random() * 100, Math.random() * 100, Red);
	gdImage.SetPixel(Math.random() * 100, Math.random() * 100, Green);
	gdImage.SetPixel(Math.random() * 100, Math.random() * 100, Blue);
}

// Draw Arc
gdImage.Arc(200, 100, 50, 50, 90, 180, Blue);
// Draw Circle
gdImage.Arc(50, 150, 50, 50, 0, 360, Red);

// Fill the Circle (With a different color)
gdImage.Fill(50, 150, Green);

// Draw Filled Arc (Can be used for Pie Charts etc.)
gdImage.FilledArc(150, 150, 50, 50, 0, 270, Red);

// Draw an Oval (Ellipse)!
gdImage.FilledEllipse(125, 100, 30, 20, Yellow);

// Draw 5 Random Lines - With Random Color! (Color gets added to the palette!)
for (var i = 0; i < 5; i++)
	gdImage.Line(Math.random() * 200, Math.random() * 200, Math.random() * 200, Math.random() * 200, gdImage.ColorAllocate(Math.random() * 255, Math.random() * 255, Math.random() * 255));

// Draw random Gradient!
gdImage.GradientFillRect(gdImage.Color(Math.random() * 255, Math.random() * 255, Math.random() * 255), gdImage.Color(Math.random() * 255, Math.random() * 255, Math.random() * 255), 50, 5, 245, 25);
gdImage.GradientFillRect(gdImage.Color(Math.random() * 255, Math.random() * 255, Math.random() * 255), gdImage.Color(Math.random() * 255, Math.random() * 255, Math.random() * 255), 225, 40, 245, 195, true);

// Draw Text
gdImage.Chars(gdImage.FontGetMediumBold(), (Math.random() * 50) + 75, 10, "Horizontal Gradient", Red);
gdImage.CharsUp(gdImage.FontGetMediumBold(), 230, 190, "Vertical Gradient", Blue);

gdImage.Chars(gdImage.FontGetMediumBold(), 25, 35, "Random", Red);
gdImage.Chars(gdImage.FontGetMediumBold(), 25, 45, "Pixels", Red);

gdImage.Chars(gdImage.FontGetMediumBold(), 185, 100, "Arc", Blue);
gdImage.Chars(gdImage.FontGetMediumBold(), 155, 130, "Pie", Red);
gdImage.Chars(gdImage.FontGetMediumBold(), 105, 75, "Elipse", Yellow);

gdImage.Chars(gdImage.FontGetTiny(), 35, 140, "Fill to", Black);
gdImage.Chars(gdImage.FontGetTiny(), 37, 150, "border", Black);

gdImage.Chars(gdImage.FontGetMediumBold(), 30, 110, "Circle", Red);

gdImage.Chars(gdImage.FontGetTiny(), 50, 185, "Programmed by Trevor Herselman!", Green);

// Before output, if we don't want a dithered Gif, run the following command!
//gdImage.TrueColorToPalette(false, 256);
// For Gradients, Dithering will help but increase the file size!

// Return a Gif data stream
Response.BinaryWrite(gdImage.ToGifStream().Read);

gdImage = null;
%>