Title: GD Library Wrapper (For ASP Developers with PHP vs ASP examples)






Online Demo:     http://www.pinnacle.co.za/gif/dynamicgif.asp

(Hit F5 to see a NEW 'randomized' gif generated!)






This wrapper was developed by Trevor Herselman. GD is copyright 2005 Boutell.Com, Inc. GD was developed by Thomas Boutell.








Files:
======

readme.txt		-- this file
bgd.dll			-- GD library binary, downloadable from www.boutell.com/gd/    NB: NOT MY WORK!!!
GDLibrary.dll		-- ActiveX wrapper DLL, developed in VB6 with full source code! Needs to be registered before use!


VB Source Code
--------------

GDLibrary.vbp		-- Visual Basic 6 project file
gdImage.cls		-- Class module file
modGDLibrary.bas	-- Standard VB module file, with VB function declarations of the bgb.dll functions
modAPI.bas		-- Standard VB module file, declaration of some Windows API functions, eg. copyMemory


ASP Examples
------------

DynamicGif.asp		-- Dynamic Gif Example
DynamicPng.asp		-- Dynamic True Color Png Example
LoadPng.asp		-- Example of loading a Png file







What is GD Library?
-------------------

Edited quote from:     www.boutell.com/gd/

"GD is an open source code library for the dynamic creation of images by programmers. GD creates PNG, JPEG and GIF images, and is commonly used to generate charts, graphics, thumbnails, graphs, text and almost anything on the fly."

FAQ:     www.boutell.com/gd/faq.html


Why another wrapper?
--------------------

Actually, it's the first of it's kind and the only wrapper available for ASP developers. GD comes standard (built-in) with PHP 4.3.x.


Why use GD?
-----------

GD is a very mature and stable project. It's unbelievably fast and easy to use! PHP developers can easily see the similarities and ASP developers will have a smoother transition to PHP by using similar code! Take advantage of tutorials written for PHP with only syntax differences!


Why use this wrapper?
---------------------

An image can be created/loaded, drawn on, saved to file or sent to a browser in less than 5 lines of code! There are several advantages ASP developers have, such as being able to read/write directly to BLOB or Image types in SQL. This wrapper creates an ADODB.Stream (ADO 2.8) object for internal use and is fully available to the client (object.Stream) with all the features and functionality that the Stream object brings!


What can GD do?
---------------

Create Pallete/True color images
GetPixel, SetPixel
Draw Lines, Rectangles, Arcs
Fill areas to borders,
Filled Arcs, Filled Ellipses
Write Horizontal/Vertical Text


What else can the wrapper do?
-----------------------------

Gradient Fills (Horizontal and Vertical)
Load/Save from/to a file, memory or database (Using ADO 2.8 Stream Object)
Output Dynamic image directly to client


Current wrapper limitations?
----------------------------

To draw various fonts, GD uses FreeType library, I'm having problems with the function declarations.
I have problems with the declaration of the variable length polygon drawing functions.









Example 1
=========


Create a 500x500 Png image and draw a Blue rectangle.



PHP
---

<?php
	$im = imagecreate(500, 500); // Create a blank 500x500 pixel image. 
	$white = imagecolorallocate($im, 255, 255, 255); // Allocate $white to the white color in $im. 
	$blue = imagecolorallocate($im, 0, 0, 255); // Allocate $blue to the blue color in $im. 
	imagerectangle($im, 3, 15, 390, 440); // Create a rectangle starting at (3, 15), the upper left corner, that goes down to (390, 440), the lower right corner. 
	header('Content-Type: image/png'); // Send the PNG content type header so the browser knows what it's getting. 
	imagepng($im); // Output the image to the browser. 
?> 



ASP (JavaScript)
----------------

<%
	var gdImage = Server.CreateObject("GDLibrary.gdImage");
	gdImage.Create(500, 500);
	gdImage.ColorAllocate(255, 255, 255);
	var Blue = gdImage.ColorAllocate(0, 0, 255);
	gdImage.Rectangle(3, 15, 390, 440, Blue);
	Response.ContentType = "image/png";
	Response.BinaryWrite(gdImage.ToPngStream().Read);
%>










Example 2
=========

Create a 230x20 image and write "My first Program with GD"



PHP
---

<?php
	header ("Content-type: image/png"); 
	$img_handle = ImageCreate (230, 20) or die ("Cannot Create image"); 
	$back_color = ImageColorAllocate ($img_handle, 0, 10, 10); 
	$txt_color = ImageColorAllocate ($img_handle, 233, 114, 191); 
	ImageString ($img_handle, 31, 5, 5,  "My first Program with GD", $txt_color); 
	ImagePng ($img_handle); 
?> 



ASP (JavaScript)
----------------

<%
	Response.ContentType = "image/png";
	var gdImage = Server.CreateObject("GDLibrary.gdImage");
	gdImage.Create(230, 20);
	gdImage.ColorAllocate(0, 10, 10);
	var TextColor = gdImage.ColorAllocate(233, 114, 191);
	gdImage.Chars(gdImage.FontGetLarge(), 5, 5, "My first Program with GD", TextColor);
	Response.BinaryWrite(gdImage.ToPngStream().Read);
%>










Example 3
=========

Draw lines 10 pixels appart in a for loop.


PHP
---

<?php 
	Header("Content-type: image/png"); 
	$height = 300; 
	$width = 300; 
	$im = ImageCreate($width, $height); 
	$bck = ImageColorAllocate($im, 10,110,100); 
	$white = ImageColorAllocate($im, 255, 255, 255); 
	ImageLine($im, 0, 0, $width, $height, $white); 
	for($i=0;$i<=299;$i=$i+10) { 
	ImageLine($im, 0, $i, $width, $height, $white); }     
	ImagePNG($im); 
?>



ASP (JavaScript)
----------------

<%
	Response.ContentType = "image/png";
	var height = 300;
	var width = 300;
	var gdImage = Server.CreateObject("GDLibrary.gdImage");
	gdImage.Create(width, height);
	gdImage.ColorAllocate(10, 110, 100);
	var white = gdImage.ColorAllocate(255, 255, 255);
	gdImage.Line(0, 0, width, height, white);
	for (var i = 0; i < width; i += 10)
		gdImage.Line(0, i, width, height, white);
	Response.BinaryWrite(gdImage.ToPngStream().Read);
%>








Example 4
=========

Draw an ellipse (oval)



PHP
---

<?php 
	Header("Content-type: image/png"); 
	$height = 300; 
	$width = 300; 
	$im = ImageCreate($width, $height); 
	$bck = ImageColorAllocate($im, 10,110,100); 
	$white = ImageColorAllocate($im, 255, 255, 255); 
	imageellipse ($im, 150, 150, 100, 200, $white); 
	ImagePNG($im); 
?> 



ASP (JavaScript)
----------------

<%
	Response.ContentType = "image/png";
	var height = 300;
	var width = 300;
	var gdImage = Server.CreateObject("GDLibrary.gdImage");
	gdImage.Create(width, height);
	gdImage.ColorAllocate(10, 110, 100);
	var white = gdImage.ColorAllocate(255, 255, 255);
	gdImage.Ellipse(150, 150, 100, 200, white);
	Response.BinaryWrite(gdImage.ToPngStream().Read);
%>









Example 5
=========

Load a Png from file, if the file is not found, create a new image with the message "Image Not Found"


PHP
---

<? 
	header ("Content-type: image/png"); 
	$im = @ImageCreateFromPNG ("php.png"); 
	if(!$im) { 
		$img_handle = ImageCreate (200, 20) or die ("Cannot Create image"); 
		$back_color = ImageColorAllocate ($img_handle, 0, 0, 0); 
		$txt_color = ImageColorAllocate ($img_handle, 255, 255, 255); 
		ImageString ($img_handle, 10, 25, 5,  "Image Not Found", $txt_color); 
		ImagePng ($img_handle); } 
	Else { 
		echo "Image is Found"; } 
?>



ASP (JavaScript)
----------------

<%
	Response.ContentType = "image/png";
	var gdImage = Server.CreateObject("GDLibrary.gdImage");
	if (!gdImage.LoadFromFile(Server.MapPath("dynamicpng.png"))) {
		gdImage.Create(200, 20);
		gdImage.ColorAllocate(0, 0, 0);
		var TextColor = gdImage.ColorAllocate(255, 255, 255);
		gdImage.Chars(gdImage.FontGetMediumBold(), 25, 5, "Image Not Found", TextColor);
		Response.BinaryWrite(gdImage.ToPngStream().Read); }
	else {
		Response.Write("Image is Found"); }
%>







This wrapper was developed by Trevor Herselman. GD is copyright 2005 Boutell.Com, Inc. GD was developed by Thomas Boutell.
