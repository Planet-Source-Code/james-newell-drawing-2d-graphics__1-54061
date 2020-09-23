VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    
    '----------------------------------------------------------------------------------------------------------------
    'Title: 2D Drawing Class
    'Description: Drawing class to draw 2D shapes and more.
    'Author: The Dude Technologies - James Newell
    'Website: www.freepgs.com/thedudetechnologies
    'Created: 26/5/2004
    'Modified: 4/6/04
    'License: Shareware... See website for details.
    '----------------------------------------------------------------------------------------------------------------
    
'apis
    'Creating rects and detecting collisions
    Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

    'object
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

    'pen
    Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    
    Public Enum PEN_STYLES
    PS_SOLID = 0
    PS_DASH = 1                    '  -------
    PS_DOT = 2                     '  .......
    PS_DASHDOT = 3                 '  _._._._
    PS_DASHDOTDOT = 4              '  _.._.._
    PS_NULL = 5
    PS_INSIDEFRAME = 6
    PS_USERSTYLE = 7
    PS_ALTERNATE = 8
    PS_STYLE_MASK = &HF
    PS_ENDCAP_ROUND = &H0
    PS_ENDCAP_SQUARE = &H100
    PS_ENDCAP_FLAT = &H200
    PS_ENDCAP_MASK = &HF00
    PS_JOIN_ROUND = &H0
    PS_JOIN_BEVEL = &H1000
    PS_JOIN_MITER = &H2000
    PS_JOIN_MASK = &HF000
    PS_COSMETIC = &H0
    PS_GEOMETRIC = &H10000
    PS_TYPE_MASK = &HF0000
    AD_COUNTERCLOCKWISE = 1
    AD_CLOCKWISE = 2
End Enum

    
    'brush
    Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
    Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
    End Type

    Public Enum BRUSH_STYLES
    BS_SOLID = 0
    BS_NULL = 1
    BS_HOLLOW = BS_NULL
    BS_HATCHED = 2
    BS_PATTERN = 3
    BS_INDEXED = 4
    BS_DIBPATTERN = 5
    BS_DIBPATTERNPT = 6
    BS_PATTERN8X8 = 7
    BS_DIBPATTERN8X8 = 8
End Enum

Public Enum HATCH_STYLES
    HS_HORIZONTAL = 0              '  -----
    HS_VERTICAL = 1                '  |||||
    HS_FDIAGONAL = 2               '  \\\\\
    HS_BDIAGONAL = 3               '  /////
    HS_CROSS = 4                   '  +++++
    HS_DIAGCROSS = 5               '  xxxxx
    HS_FDIAGONAL1 = 6
    HS_BDIAGONAL1 = 7
    HS_SOLID = 8
    HS_DENSE1 = 9
    HS_DENSE2 = 10
    HS_DENSE3 = 11
    HS_DENSE4 = 12
    HS_DENSE5 = 13
    HS_DENSE6 = 14
    HS_DENSE7 = 15
    HS_DENSE8 = 16
    HS_NOSHADE = 17
    HS_HALFTONE = 18
    HS_SOLIDCLR = 19
    HS_DITHEREDCLR = 20
    HS_SOLIDTEXTCLR = 21
    HS_DITHEREDTEXTCLR = 22
    HS_SOLIDBKCLR = 23
    HS_DITHEREDBKCLR = 24
    HS_API_MAX = 25
End Enum

    'Font
    Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
    Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
    Private Const OUT_DEFAULT_PRECIS = 0
    Private Const CLIP_DEFAULT_PRECIS = 0
    Private Const DEFAULT_QUALITY = 0
    Private Const FF_DONTCARE = 0
    Private Const DEFAULT_PITCH = 0

    Public Enum FONT_FORMAT
        DT_TOP = &H0
        DT_LEFT = &H0
        DT_CENTER = &H1
        DT_RIGHT = &H2
        DT_VCENTER = &H4
        DT_BOTTOM = &H8
        DT_WORDBREAK = &H10
        DT_SINGLELINE = &H20
        DT_EXPANDTABS = &H40
        DT_TABSTOP = &H80
        DT_NOCLIP = &H100
        DT_EXTERNALLEADING = &H200
        DT_CALCRECT = &H400
        DT_NOPREFIX = &H800
        DT_INTERNAL = &H1000
    End Enum

    'font Height
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Const LOGPIXELSY = 90


    'drawing functions
    Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
    Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    
    'back colour
    Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    Private Const TRANSPARENT = 1

'Local variable(s) to hold property value(s)
    Private mvarDC As Long 'local copy
    Private mvarFont As StdFont 'local copy
    Private mvarFontColour As Long 'local copy
    Private mvarFillColour As Long 'local copy
    Private mvarFillStyle As Long 'local copy
    Private mvarFillHatch As Long 'local copy
    Private mvarOutlineColour As Long 'local copy
    Private mvarOutlineStyle As Long 'local copy
    Private mvarOutlineWidth As Long 'local copy
    Private mvarFontFormat As Long 'local copy
    Private mvarFontAngle As Long 'local copy

    Private lOriginalPen As Long
    Private lOriginalBrush As Long
    Private lOriginalFont As Long
    Private lOriginalFontColour As Long
    Private lOriginalBkMode As Long
    
Public Property Let FontAngle(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontAngle = 5
    mvarFontAngle = vData
End Property


Public Property Get FontAngle() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontAngle
    FontAngle = mvarFontAngle
End Property



Friend Property Let FontFormat(ByVal vData As FONT_FORMAT)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontFormat = 5
    mvarFontFormat = vData
End Property


Friend Property Get FontFormat() As FONT_FORMAT
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontFormat
    FontFormat = mvarFontFormat
End Property



    
Friend Function isRectinRect(Rect1 As RECT, Rect2 As RECT) As Boolean
    Dim lResult As Long
    Dim tmpRect As RECT ' Overlapping area is placed in here
        'check to see whether specified rec is in the specified rect
            lResult = IntersectRect(tmpRect, Rect1, Rect2)
            
            If lResult <> 0 Then
                isRectinRect = True
            Else
                isRectinRect = False
            End If
End Function

Friend Function isPtinRect(Point As POINTAPI, RECT As RECT) As Boolean
    Dim lResult As Long
        'check to see whether specified pt is in the specified rect
            lResult = PtInRect(RECT, Point.x, Point.y)
            
            If lResult <> 0 Then
                isPtinRect = True
            Else
                isPtinRect = False
            End If

End Function

Friend Function createPoint(ByVal Left As Long, ByVal Top As Long) As POINTAPI
        'load info into a pointapi structure
        With createPoint
            .x = Left
            .y = Top
        End With
End Function

Friend Function createRect(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As RECT
    Dim lResult As Long
        'load info into a rect structure
            lResult = SetRect(createRect, Left, Top, Right, Bottom)
End Function

Private Sub removeStyle()
    Dim lResult As Long
        
        'Replace new objects with original so windows clears memory properly
            
            If lOriginalPen = 0 Or lOriginalBrush = 0 Or lOriginalFont = 0 Then Exit Sub
            
            'Pen
            lResult = SelectObject(mvarDC, lOriginalPen)
            lResult = DeleteObject(lResult)
            'Brush
            lResult = SelectObject(mvarDC, lOriginalBrush)
            lResult = DeleteObject(lResult)
            'Font
            lResult = SelectObject(mvarDC, lOriginalFont)
            lResult = DeleteObject(lResult)
            lResult = SetTextColor(mvarDC, lOriginalFontColour)
            lResult = SetBkMode(mvarDC, lOriginalBkMode)
End Sub

Private Sub setStyle()
    Dim lResult As Long
    Dim lPen As Long
    Dim lBrush As Long
    Dim lFont As Long
    Dim bBrush As LOGBRUSH
    
        'Set pen style
            'Create a pen
            lPen = CreatePen(mvarOutlineStyle, mvarOutlineWidth, mvarOutlineColour)
            'Copy pen onto the dc
            lOriginalPen = SelectObject(mvarDC, lPen)
        
        'Set brush style
            With bBrush
                .lbColor = mvarFillColour
                .lbStyle = mvarFillStyle
                .lbHatch = mvarFillHatch
            End With
            
            'Create a brush
            lBrush = CreateBrushIndirect(bBrush)
            'Copy brush onto the dc
            lOriginalBrush = SelectObject(mvarDC, lBrush)
            
        'Set font style
            'Create a font
            lFont = CreateFont(-MulDiv(mvarFont.Size, GetDeviceCaps(GetDC(0), LOGPIXELSY), 72), 0, mvarFontAngle * -10, mvarFontAngle * -10, mvarFont.Weight, mvarFont.Italic, mvarFont.Underline, mvarFont.Strikethrough, mvarFont.Charset, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, mvarFont.Name & Chr(0))
            'Copy pen onto the dc
            lOriginalFont = SelectObject(mvarDC, lFont)
            
            lOriginalFontColour = SetTextColor(mvarDC, mvarFontColour)
            
            lOriginalBkMode = SetBkMode(mvarDC, TRANSPARENT)
End Sub

Friend Sub drawString(Text As String, rRect As RECT)
    Dim lResult As Long
            
            'set the current styles
            setStyle
        
        'draw the text
            lResult = DrawText(mvarDC, Text, Len(Text), rRect, mvarFontFormat)
            
            'remove styles and memory
            removeStyle
        
End Sub

Friend Sub drawCircle(rRect As RECT)
    Dim lResult As Long
            
            'set the current styles
            setStyle
        
        'draw the text
            lResult = Ellipse(mvarDC, rRect.Left, rRect.Top, rRect.Right, rRect.Bottom)
            
            'remove styles and memory
            removeStyle
        
End Sub

Friend Sub drawRectangle(rRect As RECT)
    Dim lResult As Long
            
            'set the current styles
            setStyle
        
        'draw the line
            lResult = Rectangle(mvarDC, rRect.Left, rRect.Top, rRect.Right, rRect.Bottom)
            
            'remove styles and memory
            removeStyle
        
End Sub

Friend Sub drawLine(rRect As RECT)
    Dim lResult As Long
    Dim pt As POINTAPI
    
            'set the current styles
            setStyle
        
        'draw the line
            lResult = MoveToEx(mvarDC, rRect.Left, rRect.Top, pt)
            lResult = LineTo(mvarDC, rRect.Right, rRect.Bottom)
            
            'remove styles and memory
            removeStyle
        
End Sub

Friend Property Let OutlineWidth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OutlineWidth = 5
    mvarOutlineWidth = vData
End Property

Friend Property Get OutlineWidth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OutlineWidth
    OutlineWidth = mvarOutlineWidth
End Property

Friend Property Let OutlineStyle(ByVal vData As PEN_STYLES)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OutlineStyle = 5
    mvarOutlineStyle = vData
End Property

Friend Property Get OutlineStyle() As PEN_STYLES
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OutlineStyle
    OutlineStyle = mvarOutlineStyle
End Property

Friend Property Let OutlineColour(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OutlineColour = 5
    mvarOutlineColour = vData
End Property

Friend Property Get OutlineColour() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OutlineColour
    OutlineColour = mvarOutlineColour
End Property

Friend Property Let FillHatch(ByVal vData As HATCH_STYLES)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FillHatch = 5
    mvarFillHatch = vData
End Property

Friend Property Get FillHatch() As HATCH_STYLES
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FillHatch
    FillHatch = mvarFillHatch
End Property

Friend Property Let FillStyle(ByVal vData As BRUSH_STYLES)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FillStyle = 5
    mvarFillStyle = vData
End Property

Friend Property Get FillStyle() As BRUSH_STYLES
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FillStyle
    FillStyle = mvarFillStyle
End Property

Friend Property Let FillColour(ByVal vData As BRUSH_STYLES)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FillColour = 5
    mvarFillColour = vData
End Property

Friend Property Get FillColour() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FillColour
    FillColour = mvarFillColour
End Property

Friend Property Let FontColour(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontColour = 5
    mvarFontColour = vData
End Property

Friend Property Get FontColour() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontColour
    FontColour = mvarFontColour
End Property

Friend Property Set Font(ByVal vData As StdFont)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Font = Form1
    Set mvarFont = vData
End Property

Friend Property Get Font() As StdFont
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Font
    Set Font = mvarFont
End Property

Friend Property Let DC(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DC = 5
    mvarDC = vData
End Property

Friend Property Get DC() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DC
    DC = mvarDC
End Property

Private Sub Class_Initialize()
        Set mvarFont = New StdFont
End Sub

Private Sub Class_Terminate()
        Set mvarFont = Nothing
End Sub
