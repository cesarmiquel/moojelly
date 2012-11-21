Attribute VB_Name = "Module1"
'ini
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'BMP
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Const HALFTONE = 4
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046

'layers
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'line gdi
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'HELP
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'admin
Public SmellyMoo As Boolean

'main form
Public Plain_Tile(256) As Boolean

'tilemap
Public TileMap_Store() As Byte

'caching
Public ROM_Data As String, Credits As String, Status_data As String, SN_cache As String
Public Tilemap_Mode As String, tilemap_start As Long

'uncompressed
Public Uncompressed_Credits As String

'info
Public ROM_Name As String, Rom_Version As Integer, Headers_Start As Long, Current_Palette(3) As Long

'stores
Public Store() As Level_Store_type, sprite_Store() As Sprite_Store_type, Bank_Store() As Banks_Type
Public Paths(1000) As Paths_type

Type Level_Store_type
    Zipped As String
    Unzipped As String
    Unzipped_Modified As String
    
    Sprites As String
    
    Block_data As String
    
    Borders As String
    Warps As String

    Map_Bank As Integer
    Map_Sub_Bank As Integer
    Block_Bank As Integer
End Type

Type Sprite_Store_type
    X As Integer
    Y As Integer
    typeA As Byte
    TypeB As Byte
End Type

Type Banks_Type
    Bank As Byte
    Last_Sub As Byte
    Level_list As String
End Type

Type Path_script_Type
    To As Integer
    Direction() As Byte
    Style() As Byte
End Type

Type Paths_type
    X As Single
    Y As Single
    Compass(3) As Path_script_Type
    Loading As Boolean
End Type



Public Function From_ini(Section As String, key As String, Optional Default = vbNullString, Optional int_ret As Boolean, Optional INI As String)
Dim success As Long, nSize As Long, ret As String, INIfile As String
On Error GoTo oops:
ret = Space$(2048): nSize = Len(ret)

If INI = vbNullString Then
    If ROM_Name = vbNullString Then INIfile = "Config" Else INIfile = ROM_Name
    INIfile = App.path & "\config\" & INIfile & ".ini"
Else
    INIfile = INI
End If

success = GetPrivateProfileString(Section, key, "<E!!>", ret, nSize, INIfile)
If success And Left$(ret, 5) <> "<E!!>" Then
    ret = Trim(Left$(ret, success))
    If InStr(1, ret, "x") And int_ret Then
        Main.Status "INI: Sorry, working on adding hex", 4
        From_ini = Val(Right$(ret, Len(ret) - 2))
    ElseIf int_ret Then
        From_ini = Val(ret)
    Else
        From_ini = ret
    End If
Else
    If Default = vbNullString Then
        If int_ret Then From_ini = 0 Else From_ini = vbNullString
    Else
        From_ini = Default
    End If
End If
Exit Function


oops:
Main.Error_Handler "INI: " & Section & "/" & key, Err.Description, Err.Number, INIfile
End Function

Public Sub General_Load_Tile(dest As Object, Location As Long, Tile As Integer, X As Integer, Y As Integer, Optional Advanced As Boolean)
Dim ByteA As Integer, ByteB As Integer, rgb_colour As Long


Displaying = Mid(ROM_Data, Location + Tile * 16, 16)
For subline = 0 To 7
    ByteA = Asc(Mid(Displaying, (subline * 2) + 1, 1))
    ByteB = Asc(Mid(Displaying, (subline * 2) + 2, 1))
            
    If Advanced Then
        If subline = 0 Then
            Tempa = ByteA: tempb = ByteB
            Plain_Tile(X + 16 * Y) = ((ByteA = 0 Or ByteA = 255) And (ByteB = 0 Or ByteB = 255))
        ElseIf Plain_Tile(X + 16 * Y) Then
            Plain_Tile(X + 16 * Y) = ((ByteA = 0 Or ByteA = 255) And (ByteB = 0 Or ByteB = 255))
            If Tempa <> ByteA Or tempb <> ByteB Then Plain_Tile(X + 16 * Y) = False
        End If
    End If
    
    For Bit = 1 To 8
        colour = 0
        bitval = (2 ^ (Bit - 1))
        bita = ByteA And bitval
        bitb = ByteB And bitval
        If bita = 0 Then colour = 1
        If bitb = 0 Then colour = colour + 2
        rgb_colour = Current_Palette(colour)
        'rgb_colour = RGB(colour * 100, colour * 100, colour * 100)
        
        'dest.PSet ((8 - Bit) + x * 8, subline + y * 8), rgb_colour
        SetPixel dest.hdc, (8 - Bit) + X * 8, subline + Y * 8, rgb_colour
    Next Bit
Next subline
'dest.Refresh
End Sub

Public Sub TransBltOverlay(dsthDC As Long, srchDC As Long, X As Integer, Y As Integer, width As Integer, height As Integer, TransColor As Long)
    Dim maskDC As Long, TempDC As Long
    Dim hMaskBmp As Long, hTempBmp As Long, OldBmp As Long
    
    'First create some DC's. These are our gateways to assosiated bitmaps in RAM
    maskDC = CreateCompatibleDC(dsthDC)
    TempDC = CreateCompatibleDC(dsthDC)
    'Then we need the bitmaps. Note that we create a monochrome bitmap here!
    'this is a trick we use for creating a mask fast enough.
    hMaskBmp = CreateBitmap(width, height, 1, 1, ByVal 0&)
    hTempBmp = CreateCompatibleBitmap(dsthDC, width, height)
    '..then we can assign the bitmaps to the DCs
    OldBmp = SelectObject(maskDC, hMaskBmp)
    DeleteObject OldBmp
    OldBmp = SelectObject(TempDC, hTempBmp)
    DeleteObject OldBmp
    'Now we can create a mask..First we set the background color to the
    'transparent color then we copy the image into the monochrome bitmap.
    'When we are done, we reset the background color of the original source.
    TransColor = SetBkColor(srchDC, TransColor)
    BitBlt maskDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
    TransColor = SetBkColor(srchDC, TransColor)
    'The first we do with the mask is to MergePaint it into the destination.
    'this will punch a WHITE hole in the background exactly were we want the
    'graphics to be painted in.
    BitBlt TempDC, 0, 0, width, height, maskDC, 0, 0, vbSrcCopy
    BitBlt dsthDC, X, Y, width, height, TempDC, 0, 0, vbMergePaint
    'Now we delete the transparent part of our source image. To do this
    'we must invert the mask and MergePaint it into the source image. the
    'transparent area will now appear as WHITE.
    BitBlt maskDC, 0, 0, width, height, maskDC, 0, 0, vbNotSrcCopy
    BitBlt TempDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
    BitBlt TempDC, 0, 0, width, height, maskDC, 0, 0, vbMergePaint
    'Both target and source are clean, all we have to do is to AND them together!
    BitBlt dsthDC, X, Y, width, height, TempDC, 0, 0, vbSrcAnd
    'Now all we have to do is to clean up after us and free system resources..
    DeleteObject (hMaskBmp)
    DeleteObject (hTempBmp)
    DeleteDC (maskDC)
    DeleteDC (TempDC)
End Sub
