Attribute VB_Name = "modGDLibrary"
Option Explicit

'' This file was writen by Trevor Herselman
'' GD Library was written by Thomas Boutell
'' This is a module file with declarations to the functions in bgd.dll


'Header("Content-Type: image/gif");
'$string=implode($argv," ");
'$myimage = ImageCreateFromGif("images/yourimage.gif");
'$black = ImageColorAllocate($myimage, 0, 0, 0);
'xxxxxxxxxxxxxxxx $width = (imagesx($myimage)-7.5*strlen($string))/2;
'xxxxxxxxxxxxxxxx ImageString($myimage, 1, $width, 20, $string, $black);
'ImageGif($myimage);
'ImageDestroy($myimage);


' Types
' =====


' gdImage(TYPE)
' typedef struct {
Public Type gdImageType
'  /* Palette-based image pixels */
'  unsigned char ** pixels;
    pixels As Long
'  int sx;
    sx As Integer
'  int sy;
    sy As Integer
'  /* These are valid in palette images only. See also
'  /* 'alpha', which appears later in the structure to
'    preserve binary backwards compatibility */
'  int colorsTotal;
    ColorsTotal As Long
'  int red[gdMaxColors];
    Red(256) As Long
'  int green[gdMaxColors];
    Green(256) As Long
'  int blue[gdMaxColors];
    Blue(256) As Long
'  int open[gdMaxColors];
    Free(256) As Long
'  /* For backwards compatibility, this is set to the
'    first palette entry with 100% transparency,
'    and is also set and reset by the
'    gdImageColorTransparent function. Newer
'    applications can allocate palette entries
'    with any desired level of transparency; however,
'    bear in mind that many viewers, notably
'    many web browsers, fail to implement
'    full alpha channel for PNG and provide
'    support for full opacity or transparency only. */
'  int transparent;
    TRANSPARENT As Long
'  int *polyInts;
    polyInts As Long
'  int polyAllocated;
    polyAllocated As Long
'  struct gdImageStruct *brush;
'  struct gdImageStruct *tile;
'  int brushColorMap[gdMaxColors];
'  int tileColorMap[gdMaxColors];
'  int styleLength;
'  int stylePos;
'  int *style;
'  int interlace;
'  /* New in 2.0: alpha channel for palettes. Note that only
'    Macintosh Internet Explorer and (possibly) Netscape 6
'    really support multiple levels of transparency in
'    palettes, to my knowledge, as of 2/15/01. Most
'    common browsers will display 100% opaque and
'    100% transparent correctly, and do something
'    unpredictable and/or undesirable for levels
'    in between. TBB */
'  int alpha[gdMaxColors];
'  /* Truecolor flag and pixels. New 2.0 fields appear here at the
'    end to minimize breakage of existing object code. */
'  int trueColor;
'  int ** tpixels;
'  /* Should alpha channel be copied, or applied, each time a
'    pixel is drawn? This applies to truecolor images only.
'    No attempt is made to alpha-blend in palette images,
'    even if semitransparent palette entries exist.
'    To do that, build your image as a truecolor image,
'    then quantize down to 8 bits. */
'  int alphaBlendingFlag;
'  /* Should the alpha channel of the image be saved? This affects
'    PNG at the moment; other future formats may also
'    have that capability. JPEG doesn't. */
'  int saveAlphaFlag;
' } gdImage;
End Type

' gdImagePtr (TYPE)

' gdIOCtx (TYPE)
'typedef struct gdIOCtx {
'  int (*getC) (struct gdIOCtx *);
'  int (*getBuf) (struct gdIOCtx *, void *, int wanted);
'  void (*putC) (struct gdIOCtx *, int);
'  int (*putBuf) (struct gdIOCtx *, const void *, int wanted);
'  /* seek must return 1 on SUCCESS, 0 on FAILURE. Unlike fseek! */
'  int (*seek) (struct gdIOCtx *, const int);
'  long (*tell) (struct gdIOCtx *);
'  void (*gd_free) (struct gdIOCtx *);
'} gdIOCtx;

' gdFont (TYPE)
'typedef struct {
Public Type gdFont
'  /* # of characters in font */
'  int nchars;
    nchars As Long
'  /* First character is numbered... (usually 32 = space) */
'  int offset;
    offset As Long
'  /* Character width and height */
'  int w;
    w As Long
'  int h;
    h As Long
'  /* Font data; array of characters, one row after another.
'    Easily included in code, also easily loaded from
'    data files. */
'  char *data;
    data As String
'} gdFont;
End Type

' gdFontPtr (TYPE)

' gdPoint (TYPE)
'typedef struct {
Public Type gdPoint
'        int x, y;
    x As Long
    y As Long
'} gdPoint, *gdPointPtr;
End Type

' gdPointPtr (TYPE)
' gdFTStringExtra (TYPE)
' gdFTStringExtraPtr (TYPE)

' gdSource (TYPE)
' typedef struct {
'        int (*source) (void *context, char *buffer, int len);
'        void *context;
' } gdSource, *gdSourcePtr;

' gdSink (TYPE)
' typedef struct {
'        int (*sink) (void *context, char *buffer, int len);
'        void *context;
' } gdSink, *gdSinkPtr;


' Image creation, destruction, loading and saving
' ===============================================


Public Declare Function gdImageCreate Lib "bgd.dll" Alias "gdImageCreate@8" (ByVal sx As Long, ByVal sy As Long) As Long
' gdImageCreate(sx, sy) (FUNCTION)
' gdImageCreate is called to create palette-based images, with no more than 256 colors. Invoke gdImageCreate with the x and y dimensions of the desired image. gdImageCreate returns a gdImagePtr to the new image, or NULL if unable to allocate the image. The image must eventually be destroyed using gdImageDestroy().

Public Declare Function gdImageCreateTrueColor Lib "bgd.dll" Alias "gdImageCreateTrueColor@8" (ByVal sx As Long, ByVal sy As Long) As Long
' gdImageCreateTrueColor(sx, sy) (FUNCTION)
' gdImageCreateTrueColor is called to create truecolor images, with an essentially unlimited number of colors.
' Invoke gdImageCreateTrueColor with the x and y dimensions of the desired image.
' gdImageCreateTrueColor returns a gdImagePtr to the new image, or NULL if unable to allocate the image.
' The image must eventually be destroyed using gdImageDestroy().
' Truecolor images are always filled with black at creation time. There is no concept of a "background" color index.

' gdImageCreateFromJpeg(FILE *in) (FUNCTION)
Public Declare Function gdImageCreateFromJpegPtr Lib "bgd.dll" Alias "gdImageCreateFromJpegPtr@8" (ByVal size As Long, data As Any) As Long
' gdImageCreateFromJpegPtr(int size, void *data) (FUNCTION)
' gdImageCreateFromJpegCtx(gdIOCtx *in) (FUNCTION)

' gdImageCreateFromJpeg is called to load truecolor images from JPEG format files.
' Invoke gdImageCreateFromJpeg with an already opened pointer to a file containing the desired image.
' gdImageCreateFromJpeg returns a gdImagePtr to the new truecolor image, or NULL if unable to load the image (most often because the file is corrupt or does not contain a JPEG image).
' gdImageCreateFromJpeg does not close the file.
' You can inspect the sx and sy members of the image to determine its size.
' The image must eventually be destroyed using gdImageDestroy().
' The returned image is always a truecolor image.
' If you already have the image file in memory, pass the size of the file and a pointer to the file's data to gdImageCreateFromJpegPtr, which is otherwise identical to gdImageCreateFromJpeg.

' gdImageCreateFromPng(FILE *in) (FUNCTION)
Public Declare Function gdImageCreateFromPngPtr Lib "bgd.dll" Alias "gdImageCreateFromPngPtr@8" (ByVal size As Long, data As Any) As Long
' gdImageCreateFromPngPtr(int size, void *data) (FUNCTION)
' gdImageCreateFromPngCtx(gdIOCtx *in) (FUNCTION)

' gdImageCreateFromPngSource(gdSourcePtr in) (FUNCTION)

' gdImageCreateFromGif(FILE *in) (FUNCTION)
Public Declare Function gdImageCreateFromGifPtr Lib "bgd.dll" Alias "gdImageCreateFromGifPtr@8" (ByVal size As Long, data As Any) As Long
' gdImageCreateFromGifPtr(int size, void *data) (FUNCTION)
' gdImageCreateFromGifCtx(gdIOCtx *in) (FUNCTION)

' gdImageCreateFromGif is called to load images from GIF format files.
' Invoke gdImageCreateFromGif with an already opened pointer to a file containing the desired image.
' gdImageCreateFromGif returns a gdImagePtr to the new image, or NULL if unable to load the image (most often because the file is corrupt or does not contain a GIF image).
' gdImageCreateFromGif does not close the file.
' You can inspect the sx and sy members of the image to determine its size.
' The image must eventually be destroyed using gdImageDestroy().
' If you already have the image file in memory, pass the size of the file and a pointer to the file's data to gdImageCreateFromGifPtr, which is otherwise identical to gdImageCreateFromGif.

' gdImageCreateFromGd(FILE *in) (FUNCTION)
Public Declare Function gdImageCreateFromGdPtr Lib "bgd.dll" Alias "gdImageCreateFromGdPtr@8" (ByVal size As Long, data As Any) As Long
' gdImageCreateFromGdPtr(int size, void *data) (FUNCTION)
' gdImageCreateFromGdCtx(gdIOCtx *in) (FUNCTION)

' gdImageCreateFromGd is called to load images from gd format files.
' Invoke gdImageCreateFromGd with an already opened pointer to a file containing the desired image in the gd file format, which is specific to gd and intended for very fast loading.
' (It is not intended for compression; for compression, use PNG or JPEG.)
' If you already have the image file in memory, pass the size of the file and a pointer to the file's data to gdImageCreateFromGdPtr, which is otherwise identical to gdImageCreateFromGd.
' gdImageCreateFromGd returns a gdImagePtr to the new image, or NULL if unable to load the image (most often because the file is corrupt or does not contain a gd format image).
' gdImageCreateFromGd does not close the file.
' You can inspect the sx and sy members of the image to determine its size.
' The image must eventually be destroyed using gdImageDestroy().

' gdImageCreateFromGd2(FILE *in) (FUNCTION)
Public Declare Function gdImageCreateFromGd2Ptr Lib "bgd.dll" Alias "gdImageCreateFromGd2Ptr@8" (ByVal size As Long, data As Any) As Long
' gdImageCreateFromGd2Ptr(int size, void *data) (FUNCTION)
' gdImageCreateFromGd2Ctx(gdIOCtx *in) (FUNCTION)

' gdImageCreateFromGd2Part(FILE *in, int srcX, int srcY, int w, int h) (FUNCTION)
Public Declare Function gdImageCreateFromGd2PartPtr Lib "bgd.dll" Alias "gdImageCreateFromGd2PartPtr@24" (ByVal size As Long, data As Any, ByVal srcX As Long, ByVal srcY As Long, ByVal w As Long, ByVal h As Long) As Long
' gdImageCreateFromGd2PartPtr(int size, void *data, int srcX, int srcY, int w, int h) (FUNCTION)
' gdImageCreateFromGd2PartCtx(gdIOCtx *in) (FUNCTION)

' gdImageCreateFromWBMP(FILE *in) (FUNCTION)
Public Declare Function gdImageCreateFromWBMPPtr Lib "bgd.dll" Alias "gdImageCreateFromWBMPPtr@8" (ByVal size As Long, data As Any) As Long
' gdImageCreateFromWBMPPtr(int size, void *data) (FUNCTION)
' gdImageCreateFromWBMPCtx(gdIOCtx *in) (FUNCTION)

' gdImageCreateFromXbm(FILE *in) (FUNCTION)
Public Declare Function gdImageCreateFromXpm Lib "bgd.dll" Alias "gdImageCreateFromXpm@4" (FileName As String) As Long
' gdImageCreateFromXpm(char *filename) (FUNCTION)

Public Declare Sub gdImageDestroy Lib "bgd.dll" Alias "gdImageDestroy@4" (ByVal gdImagePtr As Long)
' gdImageDestroy(gdImagePtr im) (FUNCTION)
' gdImageDestroy is used to free the memory associated with an image.
' It is important to invoke gdImageDestroy before exiting your program or assigning a new image to a gdImagePtr variable.

' void gdImageJpeg(gdImagePtr im, FILE *out, int quality) (FUNCTION)
' void gdImageJpegCtx(gdImagePtr im, gdIOCtx *out, int quality) (FUNCTION)
' gdImageJpeg outputs the specified image to the specified file in JPEG format.
' The file must be open for writing.
' Under MSDOS and all versions of Windows, it is important to use "wb" as opposed to simply "w" as the mode when opening the file, and under Unix there is no penalty for doing so.
' gdImageJpeg does not close the file; your code must do so.
' If quality is negative, the default IJG JPEG quality value (which should yield a good general quality / size tradeoff for most situations) is used.
' Otherwise, for practical purposes, quality should be a value in the range 0-95, higher quality values usually implying both higher quality and larger image sizes.
' If you have set image interlacing using gdImageInterlace, this function will interpret that to mean you wish to output a progressive JPEG.
' Some programs (e.g., Web browsers) can display progressive JPEGs incrementally; this can be useful when browsing over a relatively slow communications link, for example.
' Progressive JPEGs can also be slightly smaller than sequential (non-progressive) JPEGs.
Public Declare Function gdImageJpegPtr Lib "bgd.dll" Alias "gdImageJpegPtr@12" (ByVal gdImagePtr As Long, size As Long, Optional ByRef quality As Long = -1) As Long
' void* gdImageJpegPtr(gdImagePtr im, int *size, int quality) (FUNCTION)
' Identical to gdImageJpeg except that it returns a pointer to a memory area with the JPEG data.
' This memory must be freed by the caller when it is no longer needed.
' The caller must invoke gdFree(), not free(), unless the caller is absolutely certain that the same implementations of malloc, free, etc. are used both at library build time and at application build time.
' The 'size' parameter receives the total size of the block of memory.

' void gdImageGif(gdImagePtr im, FILE *out)
' void gdImageGifCtx(gdImagePtr im, gdIOCtx *out) (FUNCTION)
Public Declare Function gdImageGifPtr Lib "bgd.dll" Alias "gdImageGifPtr@8" (ByVal gdImagePtr As Long, size As Long) As Long
' void* gdImageGifPtr(gdImagePtr im, int *size) (FUNCTION)
' Identical to gdImageGif except that it returns a pointer to a memory area with the GIF data.
' This memory must be freed by the caller when it is no longer needed.
' The caller must invoke gdFree(), not free(), unless the caller is absolutely certain that the same implementations of malloc, free, etc. are used both at library build time and at application build time.
' The 'size' parameter receives the total size of the block of memory.

' void gdImageGifAnimBegin(gdImagePtr im, FILE *out, int GlobalCM, int Loops)
' void gdImageGifAnimBeginCtx(gdImagePtr im, gdIOCtx *out, int GlobalCM, int Loops) (FUNCTION)
Public Declare Function gdImageGifAnimBeginPtr Lib "bgd.dll" Alias "gdImageGifAnimBeginPtr@16" (ByVal gdImagePtr As Long, size As Long, ByVal GlobalCM As Long, ByVal loops As Long) As Long
' void* gdImageGifAnimBeginPtr(gdImagePtr im, int *size, int GlobalCM, int Loops) (FUNCTION)
' void gdImageGifAnimAdd(gdImagePtr im, FILE *out, int LocalCM, int LeftOfs, int TopOfs, int Delay, int Disposal, gdImagePtr previm)
' void gdImageGifAnimAddCtx(gdImagePtr im, gdIOCtx *out, int LocalCM, int LeftOfs, int TopOfs, int Delay, int Disposal, gdImagePtr previm) (FUNCTION)
Public Declare Function gdImageGifAnimAddPtr Lib "bgd.dll" Alias "gdImageGifAnimAddPtr@32" (ByVal gdImagePtr As Long, size As Long, ByVal LocalCM As Long, ByVal LeftOfs As Long, ByVal TopOfs As Long, ByVal Delay As Long, ByVal Disposal As Long, ByVal gdPrevImagePtr As Long) As Long
' void* gdImageGifAnimAddPtr(gdImagePtr im, int *size, int LocalCM, int LeftOfs, int TopOfs, int Delay, int Disposal, gdImagePtr previm) (FUNCTION)
' void gdImageGifAnimEnd(FILE *out)
' void gdImageGifAnimEndCtx(gdIOCtx *out) (FUNCTION)
Public Declare Function gdImageGifAnimEndPtr Lib "bgd.dll" Alias "gdImageGifAnimEndPtr@4" (size As Long) As Long
' void* gdImageGifAnimEndPtr(int *size) (FUNCTION)

' void gdImagePng(gdImagePtr im, FILE *out)
' void gdImagePngCtx(gdImagePtr im, gdIOCtx *out) (FUNCTION)
' gdImagePng outputs the specified image to the specified file in PNG format.
' The file must be open for writing.
' Under MSDOS and all versions of Windows, it is important to use "wb" as opposed to simply "w" as the mode when opening the file, and under Unix there is no penalty for doing so.
' gdImagePng does not close the file; your code must do so.

' void gdImagePngEx(gdImagePtr im, FILE *out, int level)
' void gdImagePngCtxEx(gdImagePtr im, gdIOCtx *out, int level) (FUNCTION)
' Like gdImagePng, gdImagePngEx outputs the specified image to the specified file in PNG format.
' In addition, gdImagePngEx allows the level of compression to be specified.
' A compression level of 0 means "no compression."
' A compression level of 1 means "compressed, but as quickly as possible."
' A compression level of 9 means "compressed as much as possible to produce the smallest possible file."
' A compression level of -1 will use the default compression level at the time zlib was compiled on your system.
Public Declare Function gdImagePngPtr Lib "bgd.dll" Alias "gdImagePngPtr@8" (ByVal gdImagePtr As Long, size As Long) As Long
' void* gdImagePngPtr(gdImagePtr im, int *size) (FUNCTION)
' Identical to gdImagePng except that it returns a pointer to a memory area with the PNG data.
' This memory must be freed by the caller when it is no longer needed.
' The caller must invoke gdFree(), not free(), unless the caller is absolutely certain that the same implementations of malloc, free, etc. are used both at library build time and at application build time.
' The 'size' parameter receives the total size of the block of memory.
Public Declare Function gdImagePngPtrEx Lib "bgd.dll" Alias "gdImagePngPtrEx@12" (ByVal gdImagePtr As Long, size As Long, Optional ByVal Level As Long = 9) As Long
' void* gdImagePngPtrEx(gdImagePtr im, int *size, int level) (FUNCTION)
' Like gdImagePngPtr, gdImagePngPtrEx returns a pointer to a PNG image in allocated memory.
' In addition, gdImagePngPtrEx allows the level of compression to be specified.
' A compression level of 0 means "no compression."
' A compression level of 1 means "compressed, but as quickly as possible."
' A compression level of 9 means "compressed as much as possible to produce the smallest possible file."
' A compression level of -1 will use the default compression level at the time zlib was compiled on your system.

' gdImagePngToSink(gdImagePtr im, gdSinkPtr out) (FUNCTION)

' void gdImageWBMP(gdImagePtr im, int fg, FILE *out)
' gdImageWBMPCtx(gdIOCtx *out) (FUNCTION)(FUNCTION)
Public Declare Function gdImageWBMPPtr Lib "bgd.dll" Alias "gdImageWBMPPtr@12" (ByVal gdImagePtr As Long, size As Long, Optional ByVal Unknown As Long = 0&) As Long
' void* gdImageWBMPPtr(gdImagePtr im, int *size) (FUNCTION)

' void gdImageGd(gdImagePtr im, FILE *out) (FUNCTION)
' gdImageGd outputs the specified image to the specified file in the gd image format.
' The file must be open for writing.
' Under MSDOS and all versions of Windows, it is important to use "wb" as opposed to simply "w" as the mode when opening the file, and under Unix there is no penalty for doing so.
' gdImagePng does not close the file; your code must do so.
' The gd image format is intended for fast reads and writes of images your program will need frequently to build other images.
' It is not a compressed format, and is not intended for general use.
Public Declare Function gdImageGdPtr Lib "bgd.dll" Alias "gdImageGdPtr@8" (ByVal gdImagePtr As Long, size As Long) As Long
' void* gdImageGdPtr(gdImagePtr im, int *size) (FUNCTION)
' Identical to gdImageGd except that it returns a pointer to a memory area with the GD data.
' This memory must be freed by the caller when it is no longer needed.
' The caller must invoke gdFree(), not free(), unless the caller is absolutely certain that the same implementations of malloc, free, etc. are used both at library build time and at application build time.
' The 'size' parameter receives the total size of the block of memory.

' void gdImageGd2(gdImagePtr im, FILE *out, int chunkSize, int fmt)
' void gdImageGd2Ctx(gdImagePtr im, gdIOCtx *out, int chunkSize, int fmt) (FUNCTION)
Public Declare Function gdImageGd2Ptr Lib "bgd.dll" Alias "gdImageGd2Ptr@16" (ByVal gdImagePtr As Long, ByVal chunkSize As Long, ByVal fmt As Long, size As Long) As Long
' void* gdImageGd2Ptr(gdImagePtr im, int chunkSize, int fmt, int *size) (FUNCTION)

Public Declare Sub gdImageTrueColorToPalette Lib "bgd.dll" Alias "gdImageTrueColorToPalette@12" (ByVal gdImagePtr As Long, ByVal ditherFlag As Long, ByVal colorsWanted As Long)
' void gdImageTrueColorToPalette(gdImagePtr im, int ditherFlag, int colorsWanted)
' gdImageTrueColorToPalette permanently converts the existing image.
Public Declare Function gdImageCreatePaletteFromTrueColor Lib "bgd.dll" Alias "gdImageCreatePaletteFromTrueColor@12" (ByVal gdImagePtr As Long, ByVal ditherFlag As Long, ByVal colorsWanted As Long) As Long
' gdImagePtr gdImageCreatePaletteFromTrueColor(gdImagePtr im, int ditherFlag, int colorsWanted) (FUNCTION)
' gdImageCreatePaletteFromTrueColor returns a new image.
' The two functions are otherwise identical.


' Drawing Functions
' =================


Public Declare Sub gdImageSetPixel Lib "bgd.dll" Alias "gdImageSetPixel@16" (ByVal gdImagePtr As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long)
' void gdImageSetPixel(gdImagePtr im, int x, int y, int color) (FUNCTION)
' gdImageSetPixel sets a pixel to a particular color index.
' Always use this function or one of the other drawing functions to access pixels; do not access the pixels of the gdImage structure directly.

Public Declare Sub gdImageLine Lib "bgd.dll" Alias "gdImageLine@24" (ByVal gdImagePtr As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As Long)
' void gdImageLine(gdImagePtr im, int x1, int y1, int x2, int y2, int color) (FUNCTION)
' gdImageLine is used to draw a line between two endpoints (x1,y1 and x2, y2).
' The line is drawn using the color index specified.
' Note that the color index can be an actual color returned by gdImageColorAllocate or one of gdStyled, gdBrushed or gdStyledBrushed.

Public Declare Sub gdImageDashedLine Lib "bgd.dll" Alias "gdImageDashedLine@24" (ByVal gdImagePtr As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As Long)
' void gdImageDashedLine(gdImagePtr im, int x1, int y1, int x2, int y2, int color) (FUNCTION)
' gdImageDashedLine is provided solely for backwards compatibility with gd 1.0. New programs should draw dashed lines using the normal gdImageLine function and the new gdImageSetStyle function.

Public Declare Sub gdImagePolygon Lib "bgd.dll" Alias "gdImagePolygon@16" (ByVal gdImagePtr As Long, gdPointPtr As Any, ByVal pointsTotal As Long, ByVal color As Long)
' void gdImagePolygon(gdImagePtr im, gdPointPtr points, int pointsTotal, int color) (FUNCTION)
' gdImagePolygon is used to draw a polygon with the verticies (at least 3) specified, using the color index specified. See also gdImageFilledPolygon.

Public Declare Sub gdImageOpenPolygon Lib "bgd.dll" Alias "gdImageOpenPolygon@16" (ByVal gdImagePtr As Long, gdPointPtr As Any, ByVal pointsTotal As Long, ByVal color As Long)
' void gdImageOpenPolygon(gdImagePtr im, gdPointPtr points, int pointsTotal, int color) (FUNCTION)
' gdImageOpenPolygon is used to draw a sequence of lines with the verticies (at least 3) specified, using the color index specified. Unlike gdImagePolygon, the enpoints of the line sequence are not connected to a closed polygon.

Public Declare Sub gdImageRectangle Lib "bgd.dll" Alias "gdImageRectangle@24" (ByVal gdImagePtr As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As Long)
' void gdImageRectangle(gdImagePtr im, int x1, int y1, int x2, int y2, int color) (FUNCTION)
' gdImageRectangle is used to draw a rectangle with the two corners (upper left first, then lower right) specified, using the color index specified.

Public Declare Sub gdImageFilledPolygon Lib "bgd.dll" Alias "gdImageFilledPolygon@16" (ByVal gdImagePtr As Long, gdPointPtr As Any, ByVal pointsTotal As Long, ByVal color As Long)
' void gdImageFilledPolygon(gdImagePtr im, gdPointPtr points, int pointsTotal, int color) (FUNCTION)
' gdImageFilledPolygon is used to fill a polygon with the verticies (at least 3) specified, using the color index specified. See also gdImagePolygon.

Public Declare Sub gdImageFilledRectangle Lib "bgd.dll" Alias "gdImageFilledRectangle@24" (ByVal gdImagePtr As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As Long)
' void gdImageFilledRectangle(gdImagePtr im, int x1, int y1, int x2, int y2, int color) (FUNCTION)

Public Declare Sub gdImageArc Lib "bgd.dll" Alias "gdImageArc@32" (ByVal gdImagePtr As Long, ByVal cx As Long, ByVal cy As Long, ByVal w As Long, ByVal h As Long, ByVal s As Long, ByVal e As Long, ByVal color As Long)
' void gdImageArc(gdImagePtr im, int cx, int cy, int w, int h, int s, int e, int color) (FUNCTION)

Public Declare Sub gdImageFilledArc Lib "bgd.dll" Alias "gdImageFilledArc@36" (ByVal gdImagePtr As Long, ByVal cx As Long, ByVal cy As Long, ByVal w As Long, ByVal h As Long, ByVal s As Long, ByVal e As Long, ByVal color As Long, ByVal syle As Long)
' void gdImageFilledArc(gdImagePtr im, int cx, int cy, int w, int h, int s, int e, int color, int style) (FUNCTION)

Public Declare Sub gdImageFilledEllipse Lib "bgd.dll" Alias "gdImageFilledEllipse@24" (ByVal gdImagePtr As Long, ByVal cx As Long, ByVal cy As Long, ByVal w As Long, ByVal h As Long, ByVal color As Long)
' void gdImageFilledEllipse(gdImagePtr im, int cx, int cy, int w, int h, int color) (FUNCTION)

Public Declare Sub gdImageFillToBorder Lib "bgd.dll" Alias "gdImageFillToBorder@20" (ByVal gdImagePtr As Long, ByVal x As Long, ByVal y As Long, ByVal border As Long, ByVal color As Long)
' void gdImageFillToBorder(gdImagePtr im, int x, int y, int border, int color) (FUNCTION)

Public Declare Sub gdImageFill Lib "bgd.dll" Alias "gdImageFill@16" (ByVal gdImagePtr As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long)
' void gdImageFill(gdImagePtr im, int x, int y, int color) (FUNCTION)

Public Declare Sub gdImageSetAntiAliased Lib "bgd.dll" Alias "gdImageSetAntiAliased@8" (ByVal gdImagePtr As Long, ByVal c As Long)
' void gdImageSetAntiAliased(gdImagePtr im, int c) (FUNCTION)

Public Declare Sub gdImageSetAntiAliasedDontBlend Lib "bgd.dll" Alias "gdImageSetAntiAliasedDontBlend@12" (ByVal gdImagePtr As Long, ByVal c As Long)
' void gdImageSetAntiAliasedDontBlend(gdImagePtr im, int c) (FUNCTION)

Public Declare Sub gdImageSetBrush Lib "bgd.dll" Alias "gdImageSetBrush@8" (ByVal gdImagePtr As Long, ByVal brush As Long)
' void gdImageSetBrush(gdImagePtr im, gdImagePtr brush) (FUNCTION)

Public Declare Sub gdImageSetTile Lib "bgd.dll" Alias "gdImageSetTile@8" (ByVal gdImagePtr As Long, ByVal tile As Long)
' void gdImageSetTile(gdImagePtr im, gdImagePtr tile) (FUNCTION)

Public Declare Sub gdImageSetStyle Lib "bgd.dll" Alias "gdImageSetStyle@12" (ByVal gdImagePtr As Long, style As Long, ByVal styleLength As Long)
' void gdImageSetStyle(gdImagePtr im, int *style, int styleLength) (FUNCTION)

Public Declare Sub gdImageSetThickness Lib "bgd.dll" Alias "gdImageSetThickness@8" (ByVal gdImagePtr As Long, ByVal thickness As Long)
' void gdImageSetThickness(gdImagePtr im, int thickness) (FUNCTION)

Public Declare Sub gdImageAlphaBlending Lib "bgd.dll" Alias "gdImageAlphaBlending@8" (ByVal gdImagePtr As Long, ByVal blending As Long)
' void gdImageAlphaBlending(gdImagePtr im, int blending) (FUNCTION)

Public Declare Sub gdImageSaveAlpha Lib "bgd.dll" Alias "gdImageSaveAlpha@8" (ByVal gdImagePtr As Long, ByVal saveFlag As Long)
' void gdImageSaveAlpha(gdImagePtr im, int saveFlag) (FUNCTION)

Public Declare Sub gdImageSetClip Lib "bgd.dll" Alias "gdImageSetClip@20" (ByVal gdImagePtr As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
' void gdImageSetClip(gdImagePtr im, int x1, int y1, int x2, int y2) (FUNCTION)

Public Declare Sub gdImageGetClip Lib "bgd.dll" Alias "gdImageGetClip@20" (ByVal gdImagePtr As Long, x1P As Long, y1P As Long, x2P As Long, y2P As Long)
' void gdImageGetClip(gdImagePtr im, int *x1P, int *y1P, int *x2P, int *y2P) (FUNCTION)


' Query Functions
' ===============

' int gdImageAlpha(gdImagePtr im, int color) (MACRO)

Public Declare Function gdImageGetPixel Lib "bgd.dll" Alias "gdImageGetPixel@12" (ByVal gdImagePtr As Long, ByVal x As Long, ByVal y As Long) As Long
' int gdImageGetPixel(gdImagePtr im, int x, int y) (FUNCTION)
' gdImageGetPixel() retrieves the color index of a particular pixel. Always use this function to query pixels; do not access the pixels of the gdImage structure directly.

Public Declare Function gdImageBoundsSafe Lib "bgd.dll" Alias "gdImageBoundsSafe@12" (ByVal gdImagePtr As Long, ByVal x As Long, ByVal y As Long) As Long
' int gdImageBoundsSafe(gdImagePtr im, int x, int y) (FUNCTION)
' gdImageBoundsSafe returns true (1) if the specified point is within the current clipping rectangle, false (0) if not.
' The clipping rectangle is set by gdImageSetClip and defaults to the entire image.
' This function is intended primarily for use by those who wish to add functions to gd.
' All of the gd drawing functions already clip safely using this function or its macro equivalent in gd.c, gdImageBoundsSafeMacro.

' int gdImageGreen(gdImagePtr im, int color) (MACRO)
' int gdImageRed(gdImagePtr im, int color) (MACRO)
' int gdImageSX(gdImagePtr im) (MACRO)
' int gdImageSY(gdImagePtr im) (MACRO)


' Fonts and text-handling functions
' =================================


' gdFontPtr gdFontGetSmall(void) (FUNCTION)
Public Declare Function gdFontGetSmall Lib "bgd.dll" Alias "gdFontGetSmall@0" () As Long
' gdFontPtr gdFontGetLarge(void) (FUNCTION)
Public Declare Function gdFontGetLarge Lib "bgd.dll" Alias "gdFontGetLarge@0" () As Long
' gdFontPtr gdFontGetMediumBold(void) (FUNCTION)
Public Declare Function gdFontGetMediumBold Lib "bgd.dll" Alias "gdFontGetMediumBold@0" () As Long
' gdFontPtr gdFontGetGiant(void) (FUNCTION)
Public Declare Function gdFontGetGiant Lib "bgd.dll" Alias "gdFontGetGiant@0" () As Long
' gdFontPtr gdFontGetTiny(void) (FUNCTION)
Public Declare Function gdFontGetTiny Lib "bgd.dll" Alias "gdFontGetTiny@0" () As Long
' void gdImageChar(gdImagePtr im, gdFontPtr font, int x, int y, int c, int color) (FUNCTION)
Public Declare Sub gdImageChar Lib "bgd.dll" Alias "gdImageChar@24" (ByVal gdImagePtr As Long, ByVal font As Long, ByVal x As Long, ByVal y As Long, ByVal Char As Long, ByVal color As Long)
' void gdImageCharUp(gdImagePtr im, gdFontPtr font, int x, int y, int c, int color) (FUNCTION)
Public Declare Sub gdImageCharUp Lib "bgd.dll" Alias "gdImageCharUp@24" (ByVal gdImagePtr As Long, ByVal font As Long, ByVal x As Long, ByVal y As Long, ByVal Char As Long, ByVal color As Long)
' void gdImageString(gdImagePtr im, gdFontPtr font, int x, int y, unsigned char *s, int color) (FUNCTION)
Public Declare Sub gdImageString Lib "bgd.dll" Alias "gdImageString@24" (ByVal gdImagePtr As Long, ByVal font As Long, ByVal x As Long, ByVal y As Long, ByVal Chars As String, ByVal color As Long)
' void gdImageString16(gdImagePtr im, gdFontPtr font, int x, int y, unsigned short *s, int color) (FUNCTION)
Public Declare Sub gdImageString16 Lib "bgd.dll" Alias "gdImageString16@24" (ByVal gdImagePtr As Long, ByVal font As Long, ByVal x As Long, ByVal y As Long, ByVal Chars16 As String, ByVal color As Long)
' void gdImageStringUp(gdImagePtr im, gdFontPtr font, int x, int y, unsigned char *s, int color) (FUNCTION)
Public Declare Sub gdImageStringUp Lib "bgd.dll" Alias "gdImageStringUp@24" (ByVal gdImagePtr As Long, ByVal font As Long, ByVal x As Long, ByVal y As Long, ByVal Chars As String, ByVal color As Long)
' void gdImageStringUp16(gdImagePtr im, gdFontPtr font, int x, int y, unsigned short *s, int color) (FUNCTION)
Public Declare Sub gdImageStringUp16 Lib "bgd.dll" Alias "gdImageStringUp16@24" (ByVal gdImagePtr As Long, ByVal font As Long, ByVal x As Long, ByVal y As Long, ByVal Chars16 As String, ByVal color As Long)
' int gdFTUseFontConfig(int flag) (FUNCTION)
'Public Declare Function gdFTUseFontConfig Lib "bgd.dll" Alias "gdFTUseFontConfig@4" (ByVal flag As Long) As Long
' char *gdImageStringFT(gdImagePtr im, int *brect, int fg, char *fontname, double ptsize, double angle, int x, int y, char *string) (FUNCTION)
'Public Declare Function gdImageStringFT Lib "bgd.dll" Alias "gdImageStringFT@44" (ByVal gdImagePtr As Long, brect() As Long, ByVal fg As Long, ByVal fontname As String, ByVal ptsize As Double, ByVal angle As Double, ByVal x As Long, ByVal y As Long, ByVal text As String) As String
' char *gdImageStringFTEx(gdImagePtr im, int *brect, int fg, char *fontname, double ptsize, double angle, int x, int y, gdFTStringExtraPtr strex) (FUNCTION)
' char *gdImageStringFTCircle(gdImagePtr im, int cx, int cy, double radius, double textRadius, double fillPortion, char *font, double points, char *top, char *bottom, int fgcolor) (FUNCTION)
' char *gdImageStringTTF(gdImagePtr im, int *brect, int fg, char *fontname, double ptsize, double angle, int x, int y, char *string) (FUNCTION)
' int gdFontCacheSetup(void) (FUNCTION)
'Public Declare Function gdFontCacheSetup Lib "bgd.dll" Alias "gdFontCacheSetup@0" () As Long
' void gdFontCacheShutdown(void) (FUNCTION)
'Public Declare Sub gdFontCacheShutdown Lib "bgd.dll" Alias "gdFontCacheShutdown@0" ()


' Color-handling functions
' ========================


Public Declare Function gdImageColorAllocate Lib "bgd.dll" Alias "gdImageColorAllocate@16" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long
' int gdImageColorAllocate(gdImagePtr im, int r, int g, int b) (FUNCTION)
' gdImageColorAllocate: gdImageColorAllocate finds the first available color index in the image specified. When creating a new palette-based image, the first time you invoke this function, you are setting the background color for that image.
' Note that gdImageColorAllocate does not check for existing colors that match your request

Public Declare Function gdImageColorAllocateAlpha Lib "bgd.dll" Alias "gdImageColorAllocateAlpha@20" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal alpha As Long) As Long
' int gdImageColorAllocateAlpha(gdImagePtr im, int r, int g, int b, int a) (FUNCTION)

Public Declare Function gdImageColorClosest Lib "bgd.dll" Alias "gdImageColorClosest@16" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long
' int gdImageColorClosest(gdImagePtr im, int r, int g, int b) (FUNCTION)

Public Declare Function gdImageColorClosestAlpha Lib "bgd.dll" Alias "gdImageColorClosestAlpha@20" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal alpha As Long) As Long
' int gdImageColorClosestAlpha(gdImagePtr im, int r, int g, int b, int a) (FUNCTION)

Public Declare Function gdImageColorClosestHWB Lib "bgd.dll" Alias "gdImageColorClosestHWB@16" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long
' int gdImageColorClosestHWB(gdImagePtr im, int r, int g, int b) (FUNCTION)

Public Declare Function gdImageColorExact Lib "bgd.dll" Alias "gdImageColorExact@16" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long
' int gdImageColorExact(gdImagePtr im, int r, int g, int b) (FUNCTION)

Public Declare Function gdImageColorExactAlpha Lib "bgd.dll" Alias "gdImageColorExactAlpha@20" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal alpha As Long) As Long
' unlisted!

Public Declare Function gdImageColorResolve Lib "bgd.dll" Alias "gdImageColorResolve@16" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long
' int gdImageColorResolve(gdImagePtr im, int r, int g, int b) (FUNCTION)

Public Declare Function gdImageColorResolveAlpha Lib "bgd.dll" Alias "gdImageColorResolveAlpha@20" (ByVal gdImagePtr As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal alpha As Long) As Long
' int gdImageColorResolveAlpha(gdImagePtr im, int r, int g, int b, int a) (FUNCTION)

' int gdImageColorsTotal(gdImagePtr im) (MACRO)
' int gdImageRed(gdImagePtr im, int c) (MACRO)
' int gdImageGreen(gdImagePtr im, int c) (MACRO)
' int gdImageBlue(gdImagePtr im, int c) (MACRO)
' int gdImageGetInterlaced(gdImagePtr im) (MACRO)
' int gdImageGetTransparent(gdImagePtr im) (MACRO)

Public Declare Sub gdImageColorDeallocate Lib "bgd.dll" Alias "gdImageColorDeallocate@8" (ByVal gdImagePtr As Long, ByVal Index As Long)
' void gdImageColorDeallocate(gdImagePtr im, int color) (FUNCTION)

Public Declare Sub gdImageColorTransparent Lib "bgd.dll" Alias "gdImageColorDeallocate@8" (ByVal gdImagePtr As Long, ByVal Index As Long)
' void gdImageColorTransparent(gdImagePtr im, int color) (FUNCTION)

' void gdImageTrueColor(int red, int green, int blue) (MACRO)
' void gdTrueColorAlpha(int red, int green, int blue, int alpha) (MACRO)


' Copying and resizing functions
' ==============================


Public Declare Sub gdImageCopy Lib "bgd.dll" Alias "gdImageCopy@32" (ByVal gdDstPtr As Long, ByVal gdSrcPtr As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal width As Long, ByVal height As Long)
' void gdImageCopy(gdImagePtr dst, gdImagePtr src, int dstX, int dstY, int srcX, int srcY, int w, int h) (FUNCTION)

Public Declare Sub gdImageCopyResized Lib "bgd.dll" Alias "gdImageCopyResized@40" (ByVal gdDstPtr As Long, ByVal gdSrcPtr As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal destW As Long, ByVal destH As Long, ByVal srcW As Long, ByVal srcH As Long)
' void gdImageCopyResized(gdImagePtr dst, gdImagePtr src, int dstX, int dstY, int srcX, int srcY, int destW, int destH, int srcW, int srcH) (FUNCTION)

Public Declare Sub gdImageCopyResampled Lib "bgd.dll" Alias "gdImageCopyResampled@40" (ByVal gdDstPtr As Long, ByVal gdSrcPtr As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal destW As Long, ByVal destH As Long, ByVal srcW As Long, ByVal srcH As Long)
' void gdImageCopyResampled(gdImagePtr dst, gdImagePtr src, int dstX, int dstY, int srcX, int srcY, int destW, int destH, int srcW, int srcH) (FUNCTION)

Public Declare Sub gdImageCopyRotated Lib "bgd.dll" Alias "gdImageCopyRotated@44" (ByVal gdDstPtr As Long, ByVal gdSrcPtr As Long, ByVal dstX As Double, ByVal dstY As Double, ByVal srcX As Long, ByVal srcY As Long, ByVal srcW As Long, ByVal srcH As Long, ByVal angle As Long)
' void gdImageCopyRotated(gdImagePtr dst, gdImagePtr src, double dstX, double dstY, int srcX, int srcY, int srcW, int srcH, int angle) (FUNCTION)

Public Declare Sub gdImageCopyMerge Lib "bgd.dll" Alias "gdImageCopyMerge@36" (ByVal gdDstPtr As Long, ByVal gdSrcPtr As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal width As Long, ByVal height As Long, ByVal pct As Long)
' void gdImageCopyMerge(gdImagePtr dst, gdImagePtr src, int dstX, int dstY, int srcX, int srcY, int w, int h, int pct) (FUNCTION)

Public Declare Sub gdImageCopyMergeGray Lib "bgd.dll" Alias "gdImageCopyMergeGray@36" (ByVal gdDstPtr As Long, ByVal gdSrcPtr As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal width As Long, ByVal height As Long, ByVal pct As Long)
' void gdImageCopyMergeGray(gdImagePtr dst, gdImagePtr src, int dstX, int dstY, int srcX, int srcY, int w, int h, int pct) (FUNCTION)

Public Declare Sub gdImagePaletteCopy Lib "bgd.dll" Alias "gdImagePaletteCopy@8" (ByVal gdDstPtr As Long, ByVal gdSrcPtr As Long)
' void gdImagePaletteCopy(gdImagePtr dst, gdImagePtr src) (FUNCTION)

Public Declare Sub gdImageSquareToCircle Lib "bgd.dll" Alias "gdImageSquareToCircle@8" (ByVal gdImagePtr As Long, ByVal radius As Long)
' void gdImageSquareToCircle(gdImagePtr im, int radius) (FUNCTION)

Public Declare Sub gdImageSharpen Lib "bgd.dll" Alias "gdImageSharpen@8" (ByVal gdImagePtr As Long, ByVal pct As Long)
' void gdImageSharpen(gdImagePtr im, int pct) (FUNCTION)


' Miscellaneous Functions
' =======================

Public Declare Function gdImageCompare Lib "bgd.dll" Alias "gdImageCompare@8" (ByVal gdImagePtr1 As Long, ByVal gdImagePtr2 As Long) As Long
' int gdImageCompare(gdImagePtr im1, gdImagePtr im2) (FUNCTION)

Public Declare Sub gdImageInterlace Lib "bgd.dll" Alias "gdImageInterlace@8" (ByVal gdImagePtr As Long, ByVal interlace As Long)
' gdImageInterlace(gdImagePtr im, int interlace) (FUNCTION)

Public Declare Sub gdFree Lib "bgd.dll" Alias "gdFree@4" (ByVal gdGifPngJpgPtr As Long)
' gdFree(void *ptr) (FUNCTION)
' provides a reliable way to free memory allocated by functions such as gdImageGifPtr which return blocks of memory.


' Constants
' =========


'#define gdAntiAliased (-7)
Public Const gdAntiAliased = -7
'#define gdBrushed (-3)
Public Const gdBrushed = -3
'#define gdMaxColors 256
Public Const gdMaxColors = 256
'#define gdStyled (-2)
Public Const gdStyled = -2
'#define gdStyledBrushed (-4)
Public Const gdStyledBrushed = -4
'#define gdDashSize 4
Public Const gdDashSize = 4
'#define gdTiled (-5)
Public Const gdTiled = -5
'#define gdTransparent (-6)
Public Const gdTransparent = -6



Public Const gdArc = 0
Public Const gdPie = gdArc
Public Const gdChord = 1
Public Const gdNoFill = 2
Public Const gdEdged = 4


' Wrapper Functions
' =================


Public Function SaveAs(ByVal gdImagePtr As Long, FileName As String) As Long
    Dim hFile As Long
    Dim gdPtr As Long
    Dim size As Long
    Dim BytesWritten As Long

    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)

    Select Case LCase(Right(FileName, Len(FileName) - InStrRev(FileName, ".") + 1))
        '' AKA Select Case Extention
        Case ".gif": gdPtr = gdImageGifPtr(gdImagePtr, size): MsgBox "gif"
        Case ".png": 'gdPtr = gdImagePngPtr(gdImagePtr, size)
        Case ".jpg": 'gdPtr = gdImageGifPtr(gdImagePtr, size)
        Case ".bmp": 'gdPtr = gdImageGifPtr(gdImagePtr, size)
    End Select

    WriteFile ByVal hFile, ByVal gdPtr, ByVal size, BytesWritten

    CloseHandle (hFile)
    gdFree (gdPtr)

    SaveAs = BytesWritten

    '' Filename = Left(Filename, InStrRev(Filename, ".") - 1)
End Function


Public Function SavePtrAs(ByVal gdPtr As Long, ByVal size As Long, FileName As String) As Long
    Dim hFile As Long
    Dim BytesWritten As Long

    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    
    'Select Case LCase(Right(Filename, Len(Filename) - InStrRev(Filename, ".") + 1))
    '    '' AKA Select Case Extention
    '    Case ".gif": gdPtr = gdImageGifPtr(gdImagePtr, size): MsgBox "gif"
    '    Case ".png": 'gdPtr = gdImagePngPtr(gdImagePtr, size)
    '    Case ".jpg": 'gdPtr = gdImageGifPtr(gdImagePtr, size)
    '    Case ".bmp": 'gdPtr = gdImageGifPtr(gdImagePtr, size)
    'End Select

    WriteFile ByVal hFile, ByVal gdPtr, ByVal size, BytesWritten
    
    CloseHandle (hFile)

    SavePtrAs = BytesWritten

    '' Filename = Left(Filename, InStrRev(Filename, ".") - 1)
End Function


' 00007518    0     1 gdAlphaBlend@8
' 0000BBC8    1     2 gdDPExtractData@8
' 00017EBC    2     3 gdFTUseFontConfig@4
' 0001679C    3     4 gdFontCacheSetup@0
' 000166F8    4     5 gdFontCacheShutdown@0
' 00015264    5     6 gdFontGetGiant@0
' 00015278    6     7 gdFontGetLarge@0
' 0001528C    7     8 gdFontGetMediumBold@0
' 000152A0    8     9 gdFontGetSmall@0
' 000152B4    9    10 gdFontGetTiny@0
' 000BE75C    A    11 gdFontGiant
' 000C6774    B    12 gdFontLarge
' 000CC28C    C    13 gdFontMediumBold
' 000D10A4    D    14 gdFontSmall
' 000D38BC    E    15 gdFontTiny
' 00018078    F    16 gdFree@4
' 000166E8   10    17 gdFreeFontCache@0
' 00002E90   11    18 gdImageAABlend@4
' 00007694   12    19 gdImageAlphaBlending@8
' 00003D18   13    20 gdImageArc@32
' 000038D8   14    21 gdImageBoundsSafe@12
' 00003924   15    22 gdImageChar@24
' 00003A08   16    23 gdImageCharUp@24
' 00001BCC   17    24 gdImageColorAllocate@16
' 00001BFC   18    25 gdImageColorAllocateAlpha@20
' 000014F4   19    26 gdImageColorClosest@16
' 00001524   1A    27 gdImageColorClosestAlpha@20
' 00001990   1B    28 gdImageColorClosestHWB@16
' 00001F8C   1C    29 gdImageColorDeallocate@8
' 00001AA8   1D    30 gdImageColorExact@16
' 00001AD8   1E    31 gdImageColorExactAlpha@20
' 00001D34   1F    32 gdImageColorResolve@16
' 00001D64   20    33 gdImageColorResolveAlpha@20
' 00001FC0   21    34 gdImageColorTransparent@8
' 000071A8   22    35 gdImageCompare@8
' 00004750   23    36 gdImageCopy@32
' 00004A78   24    37 gdImageCopyMerge@36
' 00004EC0   25    38 gdImageCopyMergeGray@36
' 00005DE8   26    39 gdImageCopyResampled@40
' 000052B0   27    40 gdImageCopyResized@40
' 00005860   28    41 gdImageCopyRotated@44
' 00001048   29    42 gdImageCreate@8
' 0000A2AC   2A    43 gdImageCreateFromGd2@4
' 0000A354   2B    44 gdImageCreateFromGd2Ctx@4
' 0000A7FC   2C    45 gdImageCreateFromGd2Part@20
' 0000A910   2D    46 gdImageCreateFromGd2PartCtx@20
' 0000A85C   2E    47 gdImageCreateFromGd2PartPtr@24
' 0000A2FC   2F    48 gdImageCreateFromGd2Ptr@8
' 00009978   30    49 gdImageCreateFromGd@4
' 00009A20   31    50 gdImageCreateFromGdCtx@4
' 000099C8   32    51 gdImageCreateFromGdPtr@8
' 0000C1E0   33    52 gdImageCreateFromGif@4
' 0000C29C   34    53 gdImageCreateFromGifCtx@4
' 0000C238   35    54 gdImageCreateFromGifPtr@8
' 0000FB50   36    55 gdImageCreateFromJpeg@4
' 00010004   37    56 gdImageCreateFromJpegCtx@4
' 0000FBA0   38    57 gdImageCreateFromJpegPtr@8
' 00010D6C   39    58 gdImageCreateFromPng@4
' 00010FC8   3A    59 gdImageCreateFromPngCtx@4
' 00010DBC   3B    60 gdImageCreateFromPngPtr@8
' 00012C68   3C    61 gdImageCreateFromPngSource@4
' 00014F20   3D    62 gdImageCreateFromWBMP@4
' 00014DD8   3E    63 gdImageCreateFromWBMPCtx@4
' 00014F70   3F    64 gdImageCreateFromWBMPPtr@8
' 00006470   40    65 gdImageCreateFromXbm@4
' 000194BC   41    66 gdImageCreateFromXpm@4
' 00014734   42    67 gdImageCreatePaletteFromTrueColor@12
' 00001250   43    68 gdImageCreateTrueColor@8
' 00003408   44    69 gdImageDashedLine@24
' 000013E0   45    70 gdImageDestroy@4
' 00004240   46    71 gdImageFill@16
' 00004010   47    72 gdImageFillToBorder@20
' 00003D50   48    73 gdImageFilledArc@36
' 00003FD8   49    74 gdImageFilledEllipse@24
' 000068B8   4A    75 gdImageFilledPolygon@16
' 000046B8   4B    76 gdImageFilledRectangle@24
' 0000B650   4C    77 gdImageGd2@16
' 0000B69C   4D    78 gdImageGd2Ptr@16
' 00009DF8   4E    79 gdImageGd@8
' 00009E40   4F    80 gdImageGdPtr@8
' 0000777C   50    81 gdImageGetClip@20
' 00002D30   51    82 gdImageGetPixel@12
' 00002DD8   52    83 gdImageGetTrueColorPixel@12
' 0000D2CC   53    84 gdImageGif@8
' 0000D718   54    85 gdImageGifAnimAdd@32
' 0000D85C   55    86 gdImageGifAnimAddCtx@32
' 0000D69C   56    87 gdImageGifAnimAddPtr@32
' 0000D45C   57    88 gdImageGifAnimBegin@16
' 0000D4C0   58    89 gdImageGifAnimBeginCtx@16
' 0000D3F0   59    90 gdImageGifAnimBeginPtr@16
' 0000DED0   5A    91 gdImageGifAnimEnd@4
' 0000DF1C   5B    92 gdImageGifAnimEndCtx@4
' 0000DEEC   5C    93 gdImageGifAnimEndPtr@4
' 0000D314   5D    94 gdImageGifCtx@8
' 0000D260   5E    95 gdImageGifPtr@8
' 00007194   5F    96 gdImageInterlace@8
' 0000F3F8   60    97 gdImageJpeg@12
' 0000F5C0   61    98 gdImageJpegCtx@12
' 0000F444   62    99 gdImageJpegPtr@12
' 00002E98   63   100 gdImageLine@24
' 00006834   64   101 gdImageOpenPolygon@16
' 00002030   65   102 gdImagePaletteCopy@8
' 00011CE0   66   103 gdImagePng@8
' 00011E08   67   104 gdImagePngCtx@8
' 00011EA4   68   105 gdImagePngCtxEx@12
' 00011C94   69   106 gdImagePngEx@12
' 00011D2C   6A   107 gdImagePngPtr@8
' 00011D98   6B   108 gdImagePngPtrEx@12
' 00012C1C   6C   109 gdImagePngToSink@8
' 000067C0   6D   110 gdImagePolygon@16
' 0000457C   6E   111 gdImageRectangle@24
' 000076A8   6F   112 gdImageSaveAlpha@8
' 00007138   70   113 gdImageSetAntiAliased@8
' 00007168   71   114 gdImageSetAntiAliasedDontBlend@12
' 00006E88   72   115 gdImageSetBrush@8
' 000076BC   73   116 gdImageSetClip@20
' 000025D8   74   117 gdImageSetPixel@16
' 00006DCC   75   118 gdImageSetStyle@12
' 00006E74   76   119 gdImageSetThickness@8
' 00006FE0   77   120 gdImageSetTile@8
' 000091B8   78   121 gdImageSharpen@8
' 00008930   79   122 gdImageSquareToCircle@8
' 00003BC4   7A   123 gdImageString16@24
' 00003AEC   7B   124 gdImageString@24
' 00016748   7C   125 gdImageStringFT@44
' 00007CA0   7D   126 gdImageStringFTCircle@60
' 000168E8   7E   127 gdImageStringFTEx@48
' 00015850   7F   128 gdImageStringTTF@44
' 00003C34   80   129 gdImageStringUp16@24
' 00003B58   81   130 gdImageStringUp@24
' 00014760   82   131 gdImageTrueColorToPalette@12
' 00014FC8   83   132 gdImageWBMP@12
' 00014CC8   84   133 gdImageWBMPCtx@12
' 00015014   85   134 gdImageWBMPPtr@12
' 0000BAE8   86   135 gdNewDynamicCtx@8
' 0000BB10   87   136 gdNewDynamicCtxEx@12
' 0000EEC8   88   137 gdNewFileCtx@4
' 0000F070   89   138 gdNewSSCtx@8




