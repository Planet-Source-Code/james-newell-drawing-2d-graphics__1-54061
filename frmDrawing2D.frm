VERSION 5.00
Begin VB.Form frmDrawing2D 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "2D Drawing Example"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDrawing2D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDrawing2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '----------------------------------------------------------------------------------------------------------------
    'Title: Graphics Classes Example
    'Description: Shows how to use clsBMP and clsDraw
    'Author: The Dude Technologies - James Newell
    'Website: www.freepgs.com/thedudetechnologies
    'Created: 30/5/2004
    'Modified: 4/6/2004
    'License: Shareware... See website for details.
    '----------------------------------------------------------------------------------------------------------------
Private backBuffer As clsBMP
Private drawObject As clsDraw
Private imgHat As clsBMP

Private Sub Form_Load()
    Dim lResult As Long
        'create backbuffer to draw onto and draw obj to draw with
        Set backBuffer = New clsBMP
        Set drawObject = New clsDraw
        
        'create actual bmp x wide by y high, (True if we want it to be black and white)
        backBuffer.CreateBitmap 300, 100
        
        'draw some shapes onto the backbuffer
        drawObject.DC = backBuffer.DC
               
        'draw purple bg colour
        drawObject.FillColour = RGB(102, 102, 153)
        drawObject.drawRectangle drawObject.createRect(0, 0, backBuffer.Width, 55)
        
        'draw 3 squares
        drawObject.FillColour = RGB(204, 204, 204)
        drawObject.drawRectangle drawObject.createRect(10, 10, 25, 25)
        drawObject.drawRectangle drawObject.createRect(20, 20, 35, 35)
        drawObject.drawRectangle drawObject.createRect(30, 30, 45, 45)
        
        'draw circle with pattern inside
        drawObject.FillStyle = BS_HATCHED
        drawObject.FillHatch = HS_CROSS
        drawObject.drawCircle drawObject.createRect(70, 60, 100, 90)

        'draw lines through circle
        drawObject.OutlineColour = vbRed
        drawObject.drawLine drawObject.createRect(70, 60, 100, 90)
        drawObject.drawLine drawObject.createRect(100, 60, 70, 90)
        
        'draw an image onto the bb with transparency
        Set imgHat = New clsBMP
        imgHat.createBitmapFromFile App.Path & "\The Dude Hat - Small.bmp", True
        imgHat.CopyBitmap backBuffer.DC, backBuffer.Width - imgHat.Width - 10, (55 / 2) - (imgHat.Height / 2), True
        imgHat.CopyandStretchBitmap backBuffer.DC, 32, 60, 32, 32, True 'draw enlarged image
        Set imgHat = Nothing
        
        'set text properties and draw some text
        '"The Dude Technologies" Text
        drawObject.FontFormat = DT_NOCLIP Or DT_SINGLELINE Or DT_LEFT Or DT_VCENTER 'controls alignment etc
        drawObject.Font.Size = 14
        drawObject.FontColour = vbBlack
        drawObject.Font.Name = "Arial Rounded MT Bold"
        drawObject.drawString "TheDude", drawObject.createRect(50, 0, backBuffer.Width, 50)
        
        drawObject.Font.Bold = False
        drawObject.FontColour = vbWhite
        drawObject.Font.Name = "Century Gothic"
        drawObject.drawString "Technologies", drawObject.createRect(132, 0, backBuffer.Width, 50)
        
        'URL Text
        drawObject.FontFormat = drawObject.FontFormat Or DT_CENTER
        drawObject.Font.Size = 7
        drawObject.FontColour = RGB(204, 204, 204)
        drawObject.drawString "www.freepgs.com/thedudetechnologies", drawObject.createRect(0, 25, backBuffer.Width, 55)
        
        'Multiline text
        drawObject.FontFormat = DT_NOCLIP Or DT_WORDBREAK
        drawObject.FontColour = vbRed
        drawObject.drawString "Here is some test shapes", drawObject.createRect(110, 55, 200, 100)
        
        'Text on an angle
        drawObject.FontFormat = DT_NOCLIP
        drawObject.FontColour = vbBlue
        drawObject.FontAngle = 20
        drawObject.drawString "Text on an angle", drawObject.createRect(220, 55, 250, 100)
        
        'copy the backbuffer onto the form
        backBuffer.CopyBitmap Me.hdc, 0, 0
        
        Me.Refresh
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'remove objects from window
        Set drawObject = Nothing
        Set backBuffer = Nothing
End Sub
