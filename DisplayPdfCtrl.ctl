VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.UserControl DisplayPdfCtrl 
   ClientHeight    =   10608
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13872
   ScaleHeight     =   10608
   ScaleWidth      =   13872
   Begin VB.CheckBox Check1 
      Caption         =   "הצג לחצנים"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   9960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "DisplayPdfCtrl.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9840
      Visible         =   0   'False
      Width           =   735
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   13875
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   960
      Picture         =   "DisplayPdfCtrl.ctx":4760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9840
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "DisplayPdfCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

   
  
    Private pdf As AcroPDFLibCtl.AcroPDF
    
    Public assutaPdfPath As String
    Public assutaMacase As String
    Public IsReadDescription  As String
    
    Public IsRead  As Boolean
    Public ShoulbBeAssuta  As Boolean
    Private visibleCheck As Boolean
    
    Public Event Moveleft(n As Integer)
    Public Event MoveRight(n As Integer)

    Public Sub Init()
        assutaPdfPath = ""
        assutaMacase = ""
        IsReadDescription = ""
        IsRead
        = False
  '       MsgBox "1"
       If Not pdf Is Nothing Then
  '     MsgBox "AA1"
            pdf.Visible = False
        End If
  '       MsgBox "ASD1"
        Check1.Visible = False
        
          
          Command1.Visible = False
        Command2.Visible = False

    End Sub
   Public Sub ShowPDF()
    On Error GoTo ERR_ShowPDF
  '        MsgBox "ASD1"
    
        Check1.Visible = True
      
  'MsgBox "ASD1"

    If assutaPdfPath <> "" And PdfExist(assutaPdfPath) = True Then
   ' MsgBox "ASD1"
       If Not pdf Is Nothing Then
            pdf.Visible = True
        End If
   '       MsgBox "AS11111D1"
        LoadPDFControl
   '      MsgBox "A333SD1"
        pdf.LoadFile (assutaPdfPath)
   '   MsgBox "AS123D1"
        IsRead = True
         Exit Sub
     
     Else
     
     If Not pdf Is Nothing Then
            pdf.Visible = False
        End If

        If assutaMacase <> "" Then
         MsgBox "Path  does not exist or is not valid"
        End If

        IsRead = False

        VisibleCheckButtons = False
        Check1.Value = 0
        
       
    End If
        Exit Sub
ERR_ShowPDF:
    IsRead = False
    MsgBox "error " & Err.Description _
    & " " & Err.Number _
    & " " & Err.Source
   End Sub
      Private Sub LoadPDFControl()
    On Error GoTo ERR_LoadPDFControl
                
                  
        Set pdf = Controls.Add("AcroPDF.PDF.1", "Test")
        pdf.Top = 0
        pdf.Height = 9720
        pdf.Left = 0
        pdf.Width = 11500
        pdf.Visible = True
        
        Exit Sub
ERR_LoadPDFControl:
    
    End Sub

    Public Property Let VisibleCheckButtons(b As Boolean)
    Check1.Visible = b
      Call Check1_Click
    End Property


Private Sub Check1_Click()

If Check1.Value = 0 Then
    Command1.Visible = False
    Command2.Visible = False
Else
    Command1.Visible = True
    Command2.Visible = True
 End If
End Sub
Private Function PdfExist(assutaPdfPath As String) As Boolean
Dim fso As FileSystemObject
Dim sFilePath As String
 
    Set fso = New FileSystemObject
    
    sFilePath = assutaPdfPath
    If fso.FileExists(sFilePath) Then
       PdfExist = True
    Else
        PdfExist = False
    End If
End Function
Private Sub Command1_Click()
RaiseEvent Moveleft(27)
End Sub

Private Sub Command2_Click()
 RaiseEvent MoveRight(27)
End Sub

