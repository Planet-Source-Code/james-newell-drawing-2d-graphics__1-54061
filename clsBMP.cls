VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    
    '----------------------------------------------------------------------------------------------------------------
    'Title: Bitmap Class
    'Description: Bitmap Class to create/load bitmaps. Also creates masks.
    'Author: The Dude Technologies - James Newell
    'Website: www.freepgs.com/thedudetechnologies
    'Created: 26/5/2004
    'Modified: 4/6/04
    'License: Shareware... See website for details.
    '----------------------------------------------------------------------------------------------------------------
    
    Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Private Const LR_LOADFROMFILE = &H10
    Private Const IMAGE_BITMAP = 0
    
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    
    Private Declare Function CreateBMP Lib "gdi32" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
    Private Type Bitmap
            bmType As Long
            bmWidth As Long
            bmHeight As Long
            bmWidthBytes As Long
            bmPlanes As Integer
            bmBitsPixel As Integer
            bmBits As Long
    End Type
    
    Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

    Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
    
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Const BITSPIXEL = 12
    
    Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
    Private Type LOGBRUSH
            lbStyle As Long
            lbColor As Long
            lbHatch As Long
    End Type
    Private Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
    End Type
    
    'local variable(s) to hold property value(s)
    Private mvarDC As Long 'local copy
    Private mvarmaskDC As Long 'local copy
    Private mvarWidth As Long 'local copy
    Private mvarHeight As Long 'local copy
    Private mvarBitsPixel As Long 'local copy
    
    Private lOriginalBMP As Long
    Private lOriginalMaskBMP As Long

Public Sub createMask()
    Dim hBitmap As Long
    Dim lTmpDC As Long
    Dim hTmpBitmap As Long
    Dim lTmpOriginalBMP As Long
    Dim lResult As Long
       
        'if no image has been made or created, no point creating a mask
        If lOriginalBMP = 0 Then Exit Sub
    
        'create a temporary dc
        lTmpDC = CreateCompatibleDC(0)
        
        'create temporary bitmap (B&W)
        hTmpBitmap = CreateBMP(mvarWidth, mvarHeight, 1, 1, ByVal 0)
        
        'put this bmp onto the dc
        lTmpOriginalBMP = SelectObject(lTmpDC, hTmpBitmap)
        
        'create mask bitmap
        hBitmap = CreateBMP(mvarWidth, mvarHeight, 1, getBitsPerPixel, ByVal 0)
        
        'put this bmp onto the dc
        lOriginalMaskBMP = SelectObject(mvarmaskDC, hBitmap)
        
        'draw the mask
        lResult = BitBlt(lTmpDC, 0, 0, mvarWidth, mvarHeight, mvarDC, 0, 0, vbNotSrcCopy)
        lResult = BitBlt(mvarmaskDC, 0, 0, mvarWidth, mvarHeight, lTmpDC, 0, 0, vbSrcCopy)
        
        'replace bmp with the old one
        lResult = SelectObject(lTmpDC, lTmpOriginalBMP)
        lResult = DeleteObject(lResult)
        lResult = DeleteObject(lTmpOriginalBMP)
        
        'delete dc
        lResult = DeleteDC(lTmpDC)
    
End Sub

Public Sub CreateBitmap(lWidth As Long, lHeight As Long, Optional BandW As Boolean = False)
        Dim hBitmap As Long
        Dim lResult As Long
        Dim lBits As Long
            
            'cleanup mem
            selectOldBitmap
            
            'determine how many bits are needed
            If BandW Then
                lBits = 1 'user is using black and white image
            Else
                lBits = getBitsPerPixel
            End If
            
            'create a new bitmap
            hBitmap = CreateBMP(lWidth, lHeight, 1, lBits, ByVal 0)
            
            'put this bmp onto the dc
            lOriginalBMP = SelectObject(mvarDC, hBitmap)
            
            'remember dimensions
            mvarWidth = lWidth
            mvarHeight = lHeight
            mvarBitsPixel = lBits
            
            'clear the image
            Clear mvarDC
                       
End Sub

Public Sub createBitmapFromFile(Filename As String, Optional CreateAMask As Boolean = False)
        Dim hBitmap As Long
        Dim lResult As Long
        Dim BMP As Bitmap
            
            'cleanup mem
            selectOldBitmap
               
            'if no file name dont load a pic
            If Filename = "" Then Exit Sub
               
             'load a bmp from a file
            hBitmap = LoadImage(App.hInstance, Filename, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
            
            'put this bmp onto the dc
            lOriginalBMP = SelectObject(mvarDC, hBitmap)
            
            'get the bitmaps size
            lResult = GetObject(hBitmap, LenB(BMP), BMP)
            mvarWidth = BMP.bmWidth
            mvarHeight = BMP.bmHeight
            mvarBitsPixel = BMP.bmBitsPixel
            
            'if the user wants a mask, make the user a mask
            If CreateAMask Then Me.createMask
            
End Sub

Public Property Get DC() As Long
    DC = mvarDC
End Property

Public Property Get maskDC() As Long
    maskDC = mvarmaskDC
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Get BitsPerPixel() As Long
    BitsPerPixel = mvarBitsPixel
End Property

Private Sub Class_Initialize()
        'create DCs
        mvarDC = CreateCompatibleDC(0)
        mvarmaskDC = CreateCompatibleDC(0)
End Sub

Private Sub Class_Terminate()
    Dim lResult As Long
        'clean up memory
        selectOldBitmap
        lResult = DeleteObject(lOriginalBMP)
    
        'deletedcs
        lResult = DeleteDC(mvarmaskDC)
        lResult = DeleteDC(mvarDC)
End Sub

Private Sub selectOldBitmap()
    Dim lResult As Long
            'relpace the bitmap with the original bmp (so memory is deleted)
            If lOriginalMaskBMP = 0 Then GoTo Bitmap
            lResult = SelectObject(mvarmaskDC, lOriginalMaskBMP)
            lResult = DeleteObject(lResult)
Bitmap:
            If lOriginalBMP = 0 Then Exit Sub
            lResult = SelectObject(mvarDC, lOriginalBMP)
            lResult = DeleteObject(lResult)
            mvarWidth = 0
            mvarHeight = 0
End Sub

Private Function getBitsPerPixel() As Long
        getBitsPerPixel = GetDeviceCaps(GetDC(GetDesktopWindow()), BITSPIXEL)
End Function

Private Sub Clear(DC As Long) 'make the background white
Dim r As RECT
Dim lb As LOGBRUSH
Dim lResult As Long
Dim lBrush As Long
Dim lOldBrush As Long

    r.Right = mvarWidth
    r.Bottom = mvarHeight

    'create a brush
    lb.lbColor = GetBkColor(DC)
    lb.lbStyle = 0
    
    'create the brush
    lBrush = CreateBrushIndirect(lb)
    
    'select the brush onto the dc
    lOldBrush = SelectObject(DC, lBrush)
    
    'fill the rect
    lResult = FillRect(DC, r, lBrush)
    
    'delete
    lResult = DeleteObject(SelectObject(DC, lOldBrush))
    lResult = DeleteObject(lBrush)
End Sub

Public Sub CopyBitmap(DestDC As Long, Left As Long, Top As Long, Optional Transparent As Boolean = False)
    Dim lResult As Long
            
            If Transparent Then
                lResult = BitBlt(DestDC, Left, Top, mvarWidth, mvarHeight, mvarmaskDC, 0, 0, vbSrcPaint)
                lResult = BitBlt(DestDC, Left, Top, mvarWidth, mvarHeight, mvarDC, 0, 0, vbSrcAnd)
            Else
                lResult = BitBlt(DestDC, Left, Top, mvarWidth, mvarHeight, mvarDC, 0, 0, vbSrcCopy)
            End If
End Sub

Public Sub CopyandStretchBitmap(DestDC As Long, Left As Long, Top As Long, newWidth As Long, newHeight As Long, Optional Transparent As Boolean = False)
    Dim lResult As Long
            
            If Transparent Then
                lResult = StretchBlt(DestDC, Left, Top, newWidth, newHeight, mvarmaskDC, 0, 0, mvarWidth, mvarHeight, vbSrcPaint)
                lResult = StretchBlt(DestDC, Left, Top, newWidth, newHeight, mvarDC, 0, 0, mvarWidth, mvarHeight, vbSrcAnd)
            Else
                lResult = StretchBlt(DestDC, Left, Top, newWidth, newHeight, mvarDC, 0, 0, mvarWidth, mvarHeight, vbSrcCopy)
            End If
End Sub
