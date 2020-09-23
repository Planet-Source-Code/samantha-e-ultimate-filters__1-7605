VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFilters 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Filters"
   ClientHeight    =   4515
   ClientLeft      =   4425
   ClientTop       =   2205
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4470
      Left            =   0
      ScaleHeight     =   294
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.PictureBox picImagem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4470
      Left            =   0
      ScaleHeight     =   294
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6045
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All types |*.*"
      FontSize        =   2,54052e-29
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSaveFile 
         Caption         =   "Save..."
      End
      Begin VB.Menu mnuSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu blur 
         Caption         =   "Blur"
         Begin VB.Menu mnuBlur 
            Caption         =   "Blur..."
         End
         Begin VB.Menu mnuBlurMore 
            Caption         =   "Blur More..."
         End
      End
      Begin VB.Menu Sharpen 
         Caption         =   "Sharpen"
         Begin VB.Menu mnuSharpen 
            Caption         =   "Sharpen..."
         End
         Begin VB.Menu mnuSharpenMore 
            Caption         =   "Sharpen More..."
         End
      End
      Begin VB.Menu mnuEmboss 
         Caption         =   "Emboss..."
      End
      Begin VB.Menu Difuse 
         Caption         =   "Difuse"
         Begin VB.Menu mnuDifuse 
            Caption         =   "Difuse..."
         End
         Begin VB.Menu mnuDifuseMore 
            Caption         =   "Difuse More..."
         End
      End
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j
Dim Imagem
Dim AletatorioX, AletatorioY
Dim Pixel&
Dim Red As Integer, Green As Integer, Blue As Integer




Private Sub mnuOpenFile_Click()
CommonDialog1.Action = 1
Imagem = CommonDialog1.FileName
If Imagem = "" Then Exit Sub
Picture1.Picture = LoadPicture(Imagem)
frmFilters.Refresh

X = Picture1.ScaleWidth
Y = Picture1.ScaleHeight
If X > 750 Or Y > 750 Then
    MsgBox "Please, choose a smaller picture"
    X = 0
    Y = 0
    Exit Sub
End If
'resizing the form....
frmFilters.Width = frmFilters.ScaleX(Picture1.Width + 6, vbPixels, vbTwips)
frmFilters.Height = frmFilters.ScaleY(Picture1.Height + 30, vbPixels, vbTwips)
frmFilters.Refresh
End Sub

Private Sub mnuSaveFile_Click()
    CommonDialog1.Action = 2
    Imagem = CommonDialog1.FileName
    SavePicture Picture1.Image, Picture1
End Sub

Private Sub mnuSharpenMore_Click()

Call loading(i, j)
    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = ImageArray(0, i, j) + 0.8 * (ImageArray(0, i, j) - ImageArray(0, i - 1, j - 1))
            Green = ImageArray(1, i, j) + 0.8 * (ImageArray(1, i, j) - ImageArray(1, i - 1, j - 1))
            Blue = ImageArray(2, i, j) + 0.8 * (ImageArray(2, i, j) - ImageArray(2, i - 1, j - 1))
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            Picture1.PSet (j, i), RGB(Red, Green, Blue)
        Next
        Picture1.Refresh
    Next
End Sub

Private Sub mnuSharpen_Click()
   
   Call loading(i, j)
    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = ImageArray(0, i, j) + 0.5 * (ImageArray(0, i, j) - ImageArray(0, i - 1, j - 1))
            Green = ImageArray(1, i, j) + 0.5 * (ImageArray(1, i, j) - ImageArray(1, i - 1, j - 1))
            Blue = ImageArray(2, i, j) + 0.5 * (ImageArray(2, i, j) - ImageArray(2, i - 1, j - 1))
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            Picture1.PSet (j, i), RGB(Red, Green, Blue)
        Next
        Picture1.Refresh
    Next
End Sub

Private Sub mnuEmboss_Click()
   
   Call loading(i, j)

    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = Abs(ImageArray(0, i, j) - ImageArray(0, i + 1, j + 1) + 128)
            Green = Abs(ImageArray(1, i, j) - ImageArray(1, i + 1, j + 1) + 128)
            Blue = Abs(ImageArray(2, i, j) - ImageArray(2, i + 1, j + 1) + 128)
            Picture1.PSet (j, i), RGB(Red, Green, Blue)
        Next
        Picture1.Refresh
Next
End Sub

Private Sub mnuDifuseMore_Click()

Call loading(i, j)
For i = 2 To Y - 2
    For j = 2 To X - 2
        AletatorioX = Rnd * 6 - 2
        AletatorioY = Rnd * 6 - 2
        Red = ImageArray(0, i + AletatorioX, j + AletatorioY)
        Green = ImageArray(1, i + AletatorioX, j + AletatorioY)
        Blue = ImageArray(2, i + AletatorioX, j + AletatorioY)
        Picture1.PSet (j, i), RGB(Red, Green, Blue)
    Next
    Picture1.Refresh
Next
End Sub

Private Sub mnuBlur_Click()
    
Call loading(i, j)
    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = ImageArray(0, i - 1, j - 1) + ImageArray(0, i - 1, j) + ImageArray(0, i - 1, j + 1) + _
            ImageArray(0, i, j - 1) + ImageArray(0, i, j) + ImageArray(0, i, j + 1) + _
            ImageArray(0, i + 1, j - 1) + ImageArray(0, i + 1, j) + ImageArray(0, i + 1, j + 1)
            
            
            
            Green = ImageArray(1, i - 1, j - 1) + ImageArray(1, i - 1, j) + ImageArray(1, i - 1, j + 1) + _
            ImageArray(1, i, j - 1) + ImageArray(1, i, j) + ImageArray(1, i, j + 1) + _
            ImageArray(1, i + 1, j - 1) + ImageArray(1, i + 1, j) + ImageArray(1, i + 1, j + 1)

            
            Blue = ImageArray(2, i - 1, j - 1) + ImageArray(2, i - 1, j) + ImageArray(2, i - 1, j + 1) + _
            ImageArray(2, i, j - 1) + ImageArray(2, i, j) + ImageArray(2, i, j + 1) + _
            ImageArray(2, i + 1, j - 1) + ImageArray(2, i + 1, j) + ImageArray(2, i + 1, j + 1)

            Picture1.PSet (j, i), RGB(Red / 9, Green / 9, Blue / 9)
        Next
    Picture1.Refresh
    Next

End Sub

Private Sub mnuDifuse_Click()

Call loading(i, j)

    For i = 2 To Y - 2
        For j = 2 To X - 2
            AletatorioX = Rnd * 3 - 2
            AletatorioY = Rnd * 3 - 2
            Red = ImageArray(0, i + AletatorioX, j + AletatorioY)
            Green = ImageArray(1, i + AletatorioX, j + AletatorioY)
            Blue = ImageArray(2, i + AletatorioX, j + AletatorioY)
            Picture1.PSet (j, i), RGB(Red, Green, Blue)
        Next
        'DoEvents
        Picture1.Refresh
    Next
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuBlurMore_Click()

Call loading(i, j)

    For i = 1 To Y - 2
        For j = 1 To X - 2
            Red = ImageArray(0, i - 2, j - 2) + ImageArray(0, i - 2, j - 1) + ImageArray(0, i - 2, j) + _
            ImageArray(0, i - 2, j + 1) + ImageArray(0, i - 2, j + 2) + ImageArray(0, i - 1, j - 2) + _
            ImageArray(0, i - 1, j - 1) + ImageArray(0, i - 1, j) + ImageArray(0, i - 1, j + 1) + _
            ImageArray(0, i - 1, j + 2) + ImageArray(0, i, j - 2) + _
            ImageArray(0, i, j - 1) + ImageArray(0, i, j) + ImageArray(0, i, j + 1) + _
            ImageArray(0, i, j + 2) + ImageArray(0, i + 1, j - 2) + _
            ImageArray(0, i + 1, j - 1) + ImageArray(0, i + 1, j) + ImageArray(0, i + 1, j + 1) + _
            ImageArray(0, i + 1, j + 2) + ImageArray(0, i + 2, j - 2) + ImageArray(0, i + 2, j - 1) + _
            ImageArray(0, i + 2, j) + ImageArray(0, i + 2, j + 1) + ImageArray(0, i + 2, j + 2)
            
            
            Green = ImageArray(1, i - 2, j - 2) + ImageArray(1, i - 2, j - 1) + ImageArray(1, i - 2, j) + _
            ImageArray(1, i - 2, j + 1) + ImageArray(1, i - 2, j + 2) + ImageArray(1, i - 1, j - 2) + _
            ImageArray(1, i - 1, j - 1) + ImageArray(1, i - 1, j) + ImageArray(1, i - 1, j + 1) + _
            ImageArray(1, i - 1, j + 2) + ImageArray(1, i, j - 2) + _
            ImageArray(1, i, j - 1) + ImageArray(1, i, j) + ImageArray(1, i, j + 1) + _
            ImageArray(1, i, j + 2) + ImageArray(1, i + 1, j - 2) + _
            ImageArray(1, i + 1, j - 1) + ImageArray(1, i + 1, j) + ImageArray(1, i + 1, j + 1) + _
            ImageArray(1, i + 1, j + 2) + ImageArray(1, i + 2, j - 2) + ImageArray(1, i + 2, j - 1) + _
            ImageArray(1, i + 2, j) + ImageArray(1, i + 2, j + 1) + ImageArray(1, i + 2, j + 2)
            
            Blue = ImageArray(2, i - 2, j - 2) + ImageArray(2, i - 2, j - 1) + ImageArray(2, i - 2, j) + _
            ImageArray(2, i - 2, j + 1) + ImageArray(2, i - 2, j + 2) + ImageArray(2, i - 1, j - 2) + _
            ImageArray(2, i - 1, j - 1) + ImageArray(2, i - 1, j) + ImageArray(2, i - 1, j + 1) + _
            ImageArray(2, i - 1, j + 2) + ImageArray(2, i, j - 2) + _
            ImageArray(2, i, j - 1) + ImageArray(2, i, j) + ImageArray(2, i, j + 1) + _
            ImageArray(2, i, j + 2) + ImageArray(2, i + 1, j - 2) + _
            ImageArray(2, i + 1, j - 1) + ImageArray(2, i + 1, j) + ImageArray(2, i + 1, j + 1) + _
            ImageArray(2, i + 1, j + 2) + ImageArray(2, i + 2, j - 2) + ImageArray(2, i + 2, j - 1) + _
            ImageArray(2, i + 2, j) + ImageArray(2, i + 2, j + 1) + ImageArray(2, i + 2, j + 2)

            Picture1.PSet (j, i), RGB(Red / 25, Green / 25, Blue / 25)
        Next
    Picture1.Refresh
    Next
End Sub
