VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Doug Gaede
'version 2.0.0
'November 3, 2000
'----------------------------------------------------
'Notes:
'See the attached README.txt file.
'----------------------------------------------------
'You are free to use, modify and distribute this code.

'This class causes a Form, PictureBox, CommandButton,
'CheckBox or OptionButton to become shaped depending on the picture
'that is assigned to the Picture property.  One color in the picture
'will become the transparent color, depending on the color values passed to ShapeMe.
'The DragForm sub allows the user to drag a form that doesn't have a title bar.

'Note that you MUST set certain properties for each object manually.
'See the notes in the code below to find out which and what values.
'I set as many as I could in code, but some can not
'because they are read-only at runtime.

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Const RGN_DIFF = 4

Dim CurRgn As Long, TempRgn As Long  ' Region variables

'For dragging the form
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'these next 2 variables are used in the ShapeMe sub, but are declared here so the code in the Initialize... subs works
Private objName As Object 'the object that will be worked on
Private lngHeight As Long, lngWidth As Long 'height and width of object

Public Sub ShapeMe(Name As Object, Color As Long)

'Name = a Form or PictureBox name.
'Color = the color to convert to transparent (easiest to use RGB function to pass in this value)

Dim x As Long, y As Long 'points on form
Dim booMiddleOfSet As Boolean 'used during the gathering of transparent points in ShapeMe
Dim colPoints As Collection 'this will hold all usrPoints
Set colPoints = New Collection
Dim Z As Variant 'used during iteration through collection
Dim lngTransY As Long 'these 3 variables hold each point that will be made transparent
Dim lngTransStartX As Long
Dim lngTransEndX As Long
Dim intStoreScaleMode As Integer 'stores the commandbutton's form's scalemode
Dim lngHDC As Long 'the hDC property of the object

'set this so it can be used in the Initialization routine
Set objName = Name

'gather all points that need to be made transparent
If TypeOf objName Is Form Then 'check to see if this is a form and use GetPixel function which is a little faster than VB's Point
    
    'don't forget to set the Picture to a picture, BorderStyle = None settings manually
    'you CAN set Picture in your code since it isn't read-only at runtime
    InitializeForFormOrPictureBox
    lngHDC = objName.hDC 'device context for object, used in GetPixel.  GetDC does not work for a form (?) so must use VB's hwnd property
    booMiddleOfSet = False
    For y = 0 To lngHeight  ' Go through each column of pixels on form
        lngTransY = y
        For x = 0 To lngWidth  ' Go through each line of pixels on form
            If GetPixel(lngHDC, x, y) = Color Then  ' If the pixel's color is the transparency color, record it
                If booMiddleOfSet = False Then
                    lngTransStartX = x
                    lngTransEndX = x
                    booMiddleOfSet = True
                Else
                    lngTransEndX = x
                End If 'booMiddleOfSet = False
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(lngTransY, lngTransStartX, lngTransEndX)
                    booMiddleOfSet = False
                End If 'booMiddleOfSet = True
            End If 'GetPixel(lngHDC, X, Y) = Color
        Next x
    Next y
    'create base region to draw to; used below
    CurRgn = CreateRectRgn(0, 0, lngWidth, lngHeight)
    
ElseIf TypeOf objName Is PictureBox Then 'if a PictureBox then use Point function; a little slower but works when GetPixel doesn't
    
    'don't forget to set the Picture to a picture, BorderStyle = None settings manually
    'you CAN set Picture in your code since it isn't read-only at runtime
    InitializeForFormOrPictureBox
    'don't need a DC since we don't use GetPixel here
    booMiddleOfSet = False
    For y = 0 To lngHeight  ' Go through each column of pixels on form
        lngTransY = y
        For x = 0 To lngWidth  ' Go through each line of pixels on form
            If objName.Point(x, y) = Color Then
                If booMiddleOfSet = False Then
                    lngTransStartX = x
                    lngTransEndX = x
                    booMiddleOfSet = True
                Else
                    lngTransEndX = x
                End If 'booMiddleOfSet = False
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(lngTransY, lngTransStartX, lngTransEndX)
                    booMiddleOfSet = False
                End If 'booMiddleOfSet = True
            End If 'Name.Point(X, Y) = Color
        Next x
    Next y
    'create base region to draw to; used below
    CurRgn = CreateRectRgn(0, 0, lngWidth, lngHeight)
    
ElseIf TypeOf objName Is CommandButton Or TypeOf objName Is OptionButton Or TypeOf objName Is CheckBox Then 'check to see if this is a button
    
    'don't forget to set Picture and DownPicture to pictures, and Style = Graphical settings manually
    'you CAN set the picture properties in your code since they aren't read-only at runtime
    'I tried moving this initialization to a separate sub, but the buttons didn't draw correctly...don't know why
    'initialization
    With objName
        intStoreScaleMode = .Parent.ScaleMode 'store it to set it back when done so you don't mess with the programmer's mind
        .Parent.ScaleMode = 3 'the button's form's scalemode must = pixel
        .Caption = "" 'you can remove this line if you really want a caption, but it does weird things
        .BackColor = Color 'necessary
        .Refresh 'necessary
        lngHDC = GetWindowDC(.hWnd) 'device context (DC) for object.
        lngHeight = .Height 'faster to use a variable
        lngWidth = .Width 'faster to use a variable
    End With
    
    booMiddleOfSet = False
    For y = 0 To lngHeight ' Go through each column of pixels on form
        lngTransY = y
        For x = 0 To lngWidth ' Go through each line of pixels on form
            If GetPixel(lngHDC, x, y) = Color Then  ' If the pixel's color is the transparency color, record it
                If booMiddleOfSet = False Then
                    lngTransStartX = x
                    lngTransEndX = x
                    booMiddleOfSet = True
                Else
                    lngTransEndX = x
                End If 'booMiddleOfSet = False
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(lngTransY, lngTransStartX, lngTransEndX)
                    booMiddleOfSet = False
                End If 'booMiddleOfSet = True
            End If 'GetPixel(lngHDC, X, Y) = Color
        Next x
    Next y
    'create base region to draw to; used below
    CurRgn = CreateRectRgn(2, 2, lngWidth - 2, lngHeight - 2)
    
Else 'not a supported object
    Err.Raise vbObjectError + 512 + 2000, "TransForm", "Must pass in the name of a Form, PictureBox, CommandButton, CheckBox or OptionButton.  TransForm ShapeMe method failed."
    Exit Sub
End If 'test for each object

'create the transparent areas
For Each Z In colPoints
    TempRgn = CreateRectRgn(Z(1), Z(0), Z(2) + 1, Z(0) + 1)  ' Create a temporary pixel region for this pixel
    CombineRgn CurRgn, CurRgn, TempRgn, RGN_DIFF  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
    DeleteObject (TempRgn)  ' Delete the temporary region and free resources
Next

SetWindowRgn objName.hWnd, CurRgn, True  ' Finally set the windows region to the final product
'I do not use DeleteObject on the CurRgn, going with the advice in Dan Appleman's book:
'once set to a Window using SetWindowRgn, do not delete the region.

ReleaseDC objName.hWnd, lngHDC 'release the DC.  Does not hurt the form even though you got its DC from VB.

'final cleanup if a commandbutton
If TypeOf objName Is CommandButton Then
    objName.Parent.ScaleMode = intStoreScaleMode 'set it back
End If

Set colPoints = Nothing

End Sub
Public Sub DragForm(Who As Form, intButton As Integer)

On Error Resume Next

If intButton = vbLeftButton Then
    'Move the borderless form...
    ReleaseCapture
    SendMessage Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End If

End Sub

Private Sub InitializeForFormOrPictureBox()
'initialization for form or picturebox
With objName
    .AutoRedraw = True 'object must have this setting
    .ScaleMode = 3 'object must have this setting
    .ClipControls = False 'supposedly improves drawing performance.  Haven't seen it but doesn't hurt.
    lngHeight = .ScaleHeight 'faster to use a variable
    lngWidth = .ScaleWidth 'faster to use a variable
End With

End Sub






