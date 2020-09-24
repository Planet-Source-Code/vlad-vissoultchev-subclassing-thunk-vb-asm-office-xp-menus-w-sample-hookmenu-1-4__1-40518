VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Outlook Bar Control"
   ClientHeight    =   4956
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6732
   LinkTopic       =   "Form1"
   ScaleHeight     =   4956
   ScaleWidth      =   6732
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2196
      Left            =   840
      TabIndex        =   0
      Top             =   756
      Width           =   2700
      ExtentX         =   4762
      ExtentY         =   3873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sHtml       As String
    Dim nFile       As Integer
    
    On Error Resume Next
    sHtml = "<html>" & vbCrLf
    sHtml = sHtml & "<style>body { font-family: verdana; font-size: 8.5 pt; }</style><body>" & vbCrLf
    sHtml = sHtml & "<table width=700 border=0 cellpadding=0 cellspacing=0><tr></td>"
    sHtml = sHtml & "<p>This is my most complete project on planet-source-code. You can download it from <a target='_blank' href='http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=36529&lngWId=1'>here</a></p>" & vbCrLf
    sHtml = sHtml & "<p>UPDATE(2002-11-05): Version 1.3 Label edit fully implemented! Brand new subclasser working ok in MS Office and VS.NET so this one should fix the problems. Out-of-focus mouse wheel support (Outlook style). OLE Drag'n'Drop of groups implemented. cButton.Key property synched with parent collection. New properties: LabelEdit, AllowGroupDrag, GroupHilightIdx. Additional background style: ucsGrdTileBitmap. Bugfixes (including cMemDC). Help file updated.</p>" & vbCrLf
    sHtml = sHtml & "<p>UPDATE(2002-08-09): Version 1.2 Automatic OLE Drag&Drop fully implemented!!! New properties: UseSystemFont, FlatScrollArrows, WrapText. Additional background styles: ucsGrdAlphaBlend, ucsGrdStretchBitmap. Bugfixes and new samples. Help file updated.</p>" & vbCrLf
    sHtml = sHtml & "<p>UPDATE(2002-07-24): Version 1.1 Help file included. Additional background options. cMemDC bugfixes. VB bugfixes: now icons can also be 256 colors and truecolor. </p>" & vbCrLf
    sHtml = sHtml & "<p>This is a fairly complete emulation of outlook bar. This control is fully customizable and can emulate both outlook xp and 2000 button bar (see 'more samples') and then goes beyond. Control customization is accessible through couple of property pages. Featured is a hierarchical model (much like CSS) for defining formats of control elements (including hover/pressed/selected formats on group/items) which can be persisted (an .obf file) and a polymorphic object model for representation of group and item buttons data. Multi-line captions, multi-line tooltips (API), large&small icon styles, single/double/fixed bordes, horizontal/vertical gradients. Help is to be done (generated:-)) shortly. OLEDrag&Drop is in its infancy but still workable. Also, here you have it: the *realtime* color picker re-submitted as part of the outlook bar property pages -- check it out it's fast! Also, check out the error handler (robust one) and the DebugMode object leak info system.</p>" & vbCrLf
    sHtml = sHtml & "<p>Has been checked on win2k for GDI leaks (win9x to be done, anyone?). This is in response to recent submissions of 'commercial quality' and 'industrial strength' software to this site. Although not complete the project could easily become commercial one. Greetings go to: Ariad Software (now www.cyotek.com), vbAccelerator.com (great inspiration), and Carles P.V. (for his controls submissions:-)). Read readme.txt for the build procedure. Please report bugs and problems and leave your votes!</p>" & vbCrLf
    sHtml = sHtml & "<p><a target='_blank' href='http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=36529&lngWId=1'><img border=0 src='http://www.exhedra.com/upload/ScreenShots/PIC20028131321105405.gif' alt='100kb animated gif' width='700' height='520'></a></p>"
    sHtml = sHtml & "</td></tr></table>"
    sHtml = sHtml & "</body></html>"
    nFile = FreeFile
    Open Environ("TEMP") & "\ob.html" For Binary As nFile
    Put nFile, , sHtml
    Close nFile
    WebBrowser1.Navigate Environ("TEMP") & "\ob.html"
    WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Width < ScaleX(750, vbPixels) Then
        Width = ScaleX(750, vbPixels)
    End If
    WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Kill Environ("TEMP") & "\ob.html"
End Sub

