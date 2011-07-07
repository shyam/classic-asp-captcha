<%@language=jscript%>
<%

var Str;

var CaptchaType=0;

CaptchaType=Request.QueryString("CaptchaType");

// CAPTCHA Types
// 1: Numeric
// 2: Alphabetic
// 3: Alphanumeric
// 4: Alphanumeric with Other Characters

Session("Captcha") = genSessionString(CaptchaType);
Str = Session("Captcha");

Response.Expires=-1;

Response.ContentType = "image/gif";

Response.Expires = -1000;
var gdImage = Server.CreateObject("GDLibrary.gdImage");
//gdImage.Create(100, 30);
gdImage.CreateTrueColor(100, 30);


var Black = gdImage.ColorAllocate(0, 0, 0); // First color added will always be background color. Palette Index = 0
var Red = gdImage.ColorAllocate(255, 0, 0); // Palette Index = 1
var Green = gdImage.ColorAllocate(0, 255, 0); // Palette Index = 2
var Blue = gdImage.ColorAllocate(0, 0, 255); // Palette Index = 3 (Image is now 4 color)
var Yellow = gdImage.ColorAllocate(255, 255, 0); // Image becomes 8 color
var Magenta = gdImage.ColorAllocate(255, 0, 255);
var MyStrColor = gdImage.ColorAllocate(211, 229, 251);



for (var i = 1; i < 5; i++)
	gdImage.Line(Math.random() * 10 * i, Math.random() * 50, Math.random() * 100, Math.random() * 10, gdImage.ColorAllocate(Math.random() * 255, Math.random() * 255, Math.random() * 255));

gdImage.Arc(0, 0, 50, 50, 90, 180, Red );

gdImage.Chars(gdImage.FontGetGiant(), 10, 10, Str, MyStrColor);

gdImage.Arc(100, 0, 50, 50, 0, 360, Magenta);

for (var i = 0; i < 100; i++) {
	gdImage.SetPixel(Math.random() * 100, Math.random() * 100, Red);
	gdImage.SetPixel(Math.random() * 100, Math.random() * 100, Green);
	gdImage.SetPixel(Math.random() * 100, Math.random() * 100, Blue);
}



Response.BinaryWrite(gdImage.ToGifStream().Read);

gdImage = null;




function genSessionString(iCaptchaType)
{
	if(iCaptchaType==1)
	{
		return RandomNumber(RandomNumber(1000000));
	}
	else if(iCaptchaType==2)
	{
		return RandomStringAplhabetic();
	}
	else if(iCaptchaType==3)
	{
		return RandomStringAlphaNumeric();
	}
	else if(iCaptchaType==4)
	{
		return RandomStringAlphaNumericwithPunctuations();
	}
	else
	{
		return RandomNumber(RandomNumber(1000000));
	}
}


function RandomNumber(maxnumber)
{
    // This function could be used to genera
    //     te random numbers with any span
    // of possible numbers. For instance, if
    //     you wanted to generate a number
    // between 0 and 100000, you would do so
    //     mething like this:
    // <form><input type="button" v
    //     alue="Click Here" onClick="RandomNumber(
    //     100000)"></form>
    //
    // Hopefully this script helps you out. 
    //     Have fun =)
    //
    // - Michael Wieck
    var number = Math.round(maxnumber * Math.random());
    return number;
}

function RandomStringAplhabetic()
{
	// http://www.mediacollege.com/internet/javascript/number/random.html
	var chars = "ABCDEFGHIJKLMNOPQRSTUVWXTZabcdefghiklmnopqrstuvwxyz";
	var string_length = 6;
	var randomstring = '';
	for (var i=0; i<string_length; i++) {
		var rnum = Math.floor(Math.random() * chars.length);
		randomstring += chars.substring(rnum,rnum+1);
	}
	return randomstring;
}

function RandomStringAlphaNumeric()
{
	// http://www.mediacollege.com/internet/javascript/number/random.html
	var chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXTZ0123456789abcdefghiklmnopqrstuvwxyz0123456789";
	var string_length = 6;
	var randomstring = '';
	for (var i=0; i<string_length; i++) {
		var rnum = Math.floor(Math.random() * chars.length);
		randomstring += chars.substring(rnum,rnum+1);
	}
	return randomstring;
}

function RandomStringAlphaNumericwithPunctuations() 
{  

	// using code from http://psacake.com/web/ei.asp
    var length=8;
    var sPassword = "";
    
    var noPunction = 0;
    var randomLength = 0;
    
    if (randomLength) { 
        length = Math.random(); 
        
        length = parseInt(length * 100);
        length = (length % 7) + 6
    }
    
    for (i=0; i < length; i++) {
        numI = getRandomNum();
        if (noPunction) { while (checkPunc(numI)) { numI = getRandomNum(); } }
        sPassword = sPassword + String.fromCharCode(numI);
    }
    
    return sPassword;
  
}

function getRandomNum() {
        
    // between 0 - 1
    var rndNum = Math.random()

    // rndNum from 0 - 1000    
    rndNum = parseInt(rndNum * 1000);

    // rndNum from 33 - 127        
    rndNum = (rndNum % 94) + 33;
            
    return rndNum;
}

function checkPunc(num) {  
    if ((num >=33) && (num <=47)) { return true; }
    if ((num >=58) && (num <=64)) { return true; }    
    if ((num >=91) && (num <=96)) { return true; }
    if ((num >=123) && (num <=126)) { return true; }
    
    return false;
}
%>