VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PdfPager"
   ClientHeight    =   7356
   ClientLeft      =   120
   ClientTop       =   768
   ClientWidth     =   10968
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7356
   ScaleWidth      =   10968
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "��ҳ�ļ�"
      Height          =   396
      Left            =   3552
      TabIndex        =   33
      Top             =   5760
      Width           =   1164
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   9216
      TabIndex        =   32
      Text            =   "0"
      Top             =   1536
      Width           =   972
   End
   Begin VB.CommandButton Command10 
      Caption         =   "˫�����"
      Height          =   396
      Left            =   3744
      TabIndex        =   28
      Top             =   2112
      Width           =   1068
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2208
      Top             =   3360
   End
   Begin VB.CommandButton Command7 
      Caption         =   "����˫���"
      Height          =   396
      Left            =   2208
      TabIndex        =   25
      Top             =   5760
      Width           =   1164
   End
   Begin VB.PictureBox Color1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9216
      ScaleHeight     =   276
      ScaleWidth      =   276
      TabIndex        =   24
      Top             =   1056
      Width           =   300
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   1728
      TabIndex        =   15
      Text            =   "200"
      Top             =   5088
      Width           =   972
   End
   Begin VB.CommandButton Command6 
      Caption         =   "���"
      Height          =   396
      Left            =   6240
      TabIndex        =   14
      Top             =   4512
      Width           =   1068
   End
   Begin VB.TextBox Text4 
      Height          =   264
      Left            =   1728
      TabIndex        =   13
      Text            =   "D:\pages.pdf"
      Top             =   4608
      Width           =   4428
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   3660
      Left            =   10368
      TabIndex        =   11
      Top             =   2400
      Width           =   204
      _ExtentX        =   360
      _ExtentY        =   6456
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      SelStart        =   92
      Value           =   94
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   204
      Left            =   7776
      TabIndex        =   10
      Top             =   5952
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   360
      _Version        =   393216
      Max             =   100
      SelStart        =   90
      Value           =   90
   End
   Begin VB.PictureBox Box1 
      BackColor       =   &H8000000B&
      Height          =   3508
      Left            =   7872
      ScaleHeight     =   3456
      ScaleWidth      =   2436
      TabIndex        =   8
      Top             =   2496
      Width           =   2480
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ѡ��"
      Height          =   300
      Left            =   9600
      TabIndex        =   7
      Top             =   1056
      Width           =   588
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   9216
      TabIndex        =   6
      Text            =   "22"
      Top             =   576
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���ɵ����"
      Height          =   396
      Left            =   864
      TabIndex        =   5
      Top             =   5760
      Width           =   1164
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   396
      Left            =   2304
      TabIndex        =   4
      Top             =   2112
      Width           =   1068
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1632
      TabIndex        =   3
      Text            =   "200"
      Top             =   1440
      Width           =   780
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   396
      Left            =   0
      TabIndex        =   2
      Top             =   6960
      Width           =   10968
      _ExtentX        =   19346
      _ExtentY        =   699
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "�ȴ���"
            TextSave        =   "�ȴ���"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   396
      Left            =   6240
      TabIndex        =   1
      Top             =   864
      Width           =   1068
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1632
      TabIndex        =   0
      Text            =   "D:\234.pdf"
      Top             =   960
      Width           =   4428
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   768
      Top             =   3264
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "ѡ��PDFĿ¼"
      Filter          =   "*.pdf|*.pdf"
      InitDir         =   "c:\"
   End
   Begin VB.Frame Frame2 
      Caption         =   "���ɱ���ģ��"
      Height          =   2508
      Left            =   192
      TabIndex        =   17
      Top             =   4032
      Width           =   7404
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   876
         Left            =   4704
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1440
         Width           =   2604
      End
      Begin VB.Label Label10 
         Caption         =   "˵��"
         Height          =   204
         Left            =   4704
         TabIndex        =   30
         Top             =   1248
         Width           =   588
      End
      Begin VB.Label Label7 
         Caption         =   "���ҳ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   384
         TabIndex        =   19
         Top             =   1056
         Width           =   1164
      End
      Begin VB.Label Label8 
         Caption         =   "����·����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   384
         TabIndex        =   18
         Top             =   576
         Width           =   1164
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�������ĵ�����"
      Height          =   2412
      Left            =   192
      TabIndex        =   20
      Top             =   384
      Width           =   7404
      Begin VB.TextBox Text6 
         Height          =   300
         Left            =   4512
         TabIndex        =   27
         Text            =   "1"
         Top             =   1056
         Width           =   780
      End
      Begin VB.Label Label9 
         Caption         =   "��ʼ���룺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3264
         TabIndex        =   26
         Top             =   1056
         Width           =   1164
      End
      Begin VB.Label Label4 
         Caption         =   "DPI��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   768
         TabIndex        =   22
         Top             =   1056
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "����PDF��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   384
         TabIndex        =   21
         Top             =   576
         Width           =   1548
      End
   End
   Begin VB.Label Label11 
      Caption         =   "��ת�Ƕȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   7968
      TabIndex        =   31
      Top             =   1536
      Width           =   1164
   End
   Begin VB.Label Label1 
      Caption         =   "������ɫ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   7968
      TabIndex        =   23
      Top             =   1056
      Width           =   1164
   End
   Begin VB.Label Label3 
      Caption         =   "�����С��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   7968
      TabIndex        =   16
      Top             =   576
      Width           =   1164
   End
   Begin VB.Label Label6 
      Caption         =   "����λ�ã�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7968
      TabIndex        =   12
      Top             =   2112
      Width           =   1164
   End
   Begin VB.Label Label5 
      Caption         =   "A4ֽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   8832
      TabIndex        =   9
      Top             =   6336
      Width           =   684
   End
   Begin VB.Menu Menu2 
      Caption         =   "����"
      Begin VB.Menu Tab1 
         Caption         =   "������"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Single
Dim f As Integer
Dim f2 As Integer
Dim pb
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub Command1_Click()
    SB1.Panels(1).Text = "������..."
    If check1() = True Then
        time = Timer
        Dim pb As New DebenuPDFLibraryAX1016.PDFLibrary
        Dim PdfFile As String
        Dim pages As Integer
        Dim SavePath As String
        Dim SaveName As String
        Dim x As Integer, y As Integer
        Dim size As Double
        Dim begin As Integer
        Dim r%, g%, b%
        
        PdfFile = Text1.Text
        size = Text3.Text
        begin = Int(Text6.Text)
        
        Call pb.UnlockKey("j87ig3k84fb9eq9dy34z7u66y")
        Call pb.LoadFromFile(PdfFile, "")
        
        If pb.PageWidth > pb.PageHeight Then
            x = pb.PageWidth * (100 - Slider1.Value) / 100
            y = pb.PageHeight * (100 - Slider2.Value) / 100
        Else
            y = pb.PageWidth * (100 - Slider1.Value) / 100
            x = pb.PageHeight * (100 - Slider2.Value) / 100
            Call pb.SetOrigin(3)        '���½�Ϊ������ԭ��
        End If
        pages = pb.PageCount()
        
        If pb.PageWidth > pb.PageHeight Then        '�����ȡ���Ŀ�ͳ��෴�����
            For i = 1 To pages
                Call pb.SelectPage(i)
                Call pb.SetFontFlags(1, 0, 1, 0, 0, 0, 0, 1)
                pb.SetTextSize size
                pb.SetTextColor r, g, b
                Call pb.DrawRotatedText(y, x, 270, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
            Next i
        Else
            For i = 1 To pages
                Call pb.SelectPage(i)
                Call pb.SetFontFlags(1, 0, 1, 0, 0, 0, 0, 1)
                pb.SetTextSize size
                pb.SetTextColor r, g, b
                Call pb.DrawText(y, x, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
            Next i
        End If
        
        SavePath = Left(PdfFile, InStrRev(PdfFile, "\"))
        SaveName = Mid(PdfFile, InStrRev(PdfFile, "\") + 1, InStrRev(PdfFile, ".") - 4) & "_paged.pdf"
        result = pb.SaveToFile(SavePath & SaveName)
        If result = 0 Then MsgBox ("����ʧ�ܣ����ȹر��ļ�")
    End If
    SB1.Panels(1).Text = "��� ��ʱ��" & Timer - time & "s"
    Timer1.Enabled = True
End Sub

Private Sub Command10_Click()
    SB1.Panels(1).Text = "������..."
    If check1() = True Then
        time = Timer
        Dim pb As New DebenuPDFLibraryAX1016.PDFLibrary
        Dim PdfFile As String
        Dim SavePath As String
        Dim SaveName As String
        Dim pages As Integer
        Dim x As Integer, y As Integer
        Dim size As Double
        Dim begin As Integer
        Dim r%, g%, b%
        Dim odd As Boolean          '��ʼ����Ϊ�������
        
        PdfFile = Text1.Text
        size = Text3.Text
        begin = Int(Text6.Text)
        
        Call pb.UnlockKey("j87ig3k84fb9eq9dy34z7u66y")
        Call pb.LoadFromFile(PdfFile, "")
        
        If pb.PageWidth > pb.PageHeight Then
            x = pb.PageWidth * (100 - Slider1.Value) / 100
            y = pb.PageHeight * (100 - Slider2.Value) / 100
        Else
            y = pb.PageWidth * (100 - Slider1.Value) / 100
            x = pb.PageHeight * (100 - Slider2.Value) / 100
            Call pb.SetOrigin(3)        '���½�Ϊ������ԭ��
        End If
        If begin Mod 2 <> 0 Then
            odd = True
        Else
            odd = False
        End If
        pages = pb.PageCount()
        
        If pb.PageWidth > pb.PageHeight Then        '�����ȡ���Ŀ�ͳ��෴�����
            If odd = True Then
                For i = 1 To pages
                    Call pb.SelectPage(i)
                    Call pb.SetFontFlags(1, 0, 1, 0, 0, 0, 0, 1)
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    If i Mod 2 <> 0 Then
                        Call pb.DrawRotatedText(y, x, 270, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    Else
                        Call pb.DrawRotatedText(y, pb.PageHeight - x + 1.5 * size, 270, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    End If
                Next i
            Else
                For i = 1 To pages
                    Call pb.SelectPage(i)
                    Call pb.SetFontFlags(1, 0, 1, 0, 0, 0, 0, 1)
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    If i Mod 2 = 0 Then
                        Call pb.DrawRotatedText(y, pb.PageHeight - x + 1.5 * size, 270, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    Else
                        Call pb.DrawRotatedText(y, x, 270, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    End If
                Next i
            End If
        Else
            If odd = True Then
                For i = 1 To pages
                    Call pb.SelectPage(i)
                    Call pb.SetFontFlags(1, 0, 1, 0, 0, 0, 0, 1)
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    If i Mod 2 <> 0 Then
                        Call pb.DrawText(y, x, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    Else
                        Call pb.DrawText(pb.PageWidth - y + 1.5 * size, x, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    End If
                Next i
            Else
                For i = 1 To pages
                    Call pb.SelectPage(i)
                    Call pb.SetFontFlags(1, 0, 1, 0, 0, 0, 0, 1)
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    If i Mod 2 = 0 Then
                        Call pb.DrawText(pb.PageWidth - y + 1.5 * size, x, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    Else
                        Call pb.DrawText(y, x, String(3 - Len(i + begin - 1), "0") & (i + begin - 1))
                    End If
                Next i
            End If
        End If
        
        SavePath = Left(PdfFile, InStrRev(PdfFile, "\"))
        SaveName = Mid(PdfFile, InStrRev(PdfFile, "\") + 1, InStrRev(PdfFile, ".") - 4) & "_paged.pdf"
        result = pb.SaveToFile(SavePath & SaveName)
        If result = 0 Then MsgBox ("����ʧ�ܣ����ȹر��ļ�")
    End If
    SB1.Panels(1).Text = "��� ��ʱ��" & Timer - time & "s"
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    cd1.ShowOpen
    Text1.Text = cd1.FileName
End Sub

Private Sub Command3_Click()
    SB1.Panels(1).Text = "������..."
    If check2() = True Then
        time = Timer
        Dim pb As New DebenuPDFLibraryAX1016.PDFLibrary
        Dim size As Double
        Dim page As Integer
        Dim maxpage As Integer
        Dim r%, g%, b%
        Dim x%, y%
        
        r = (cd1.Color And &HFF) Mod 256
        g = ((cd1.Color And &HFF00) \ &H100) Mod 256
        b = ((cd1.Color And &HFF0000) \ &H10000) Mod 256
        Call pb.UnlockKey("j87ig3k84fb9eq9dy34z7u66y")
        
        size = Val(Text3.Text)
        angle = Val(Text8.Text)
        pb.SetOrigin 1
        maxpage = Val(Text5.Text)
        page = 1
        
        pb.SelectPage page     '��һҳ
        pb.SetPageSize "A4"
        x = pb.PageWidth * Slider1.Value / 100
        y = pb.PageHeight * Slider2.Value / 100
        pb.SetTextSize size
        pb.SetTextColor r, g, b
        Call pb.DrawRotatedText(x, y, angle, String(4 - Len(Str(page)), "0") & page)
        page = page + 1
        
        Do While page <= maxpage
            Call pb.NewPage
            pb.SetPageSize "A4"
            pb.SetTextSize size
            pb.SetTextColor r, g, b
            Call pb.DrawRotatedText(x, y, angle, String(4 - Len(Str(page)), "0") & page)
            page = page + 1
        Loop
            
        result = pb.SaveToFile(Text4.Text)
        If result = 0 Then MsgBox ("����ʧ�ܣ����ȹر��ļ�")
    End If
    SB1.Panels(1).Text = "��� ��ʱ��" & Timer - time & "s"
    Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
    cd1.ShowColor
    Color1.BackColor = cd1.Color
End Sub

Private Sub Command5_Click()
    SB1.Panels(1).Text = "������..."
    If check2() = True Then
        time = Timer
        Dim pb As New DebenuPDFLibraryAX1016.PDFLibrary
        Dim size As Double
        Dim page As Integer
        Dim maxpage As Integer
        Dim r%, g%, b%
        Dim x%, y%
        
        r = (cd1.Color And &HFF) Mod 256
        g = ((cd1.Color And &HFF00) \ &H100) Mod 256
        b = ((cd1.Color And &HFF0000) \ &H10000) Mod 256
        Call pb.UnlockKey("j87ig3k84fb9eq9dy34z7u66y")
        
        size = Val(Text3.Text)
        angle = Val(Text8.Text)
        pb.SetOrigin 1
        maxpage = Val(Text5.Text)
        page = 1
    
        pb.SelectPage page     '��һҳ
        pb.SetPageSize "A4"
        x = pb.PageWidth * Slider1.Value / 100
        y = pb.PageHeight * Slider2.Value / 100
        pb.SetTextSize size
        pb.SetTextColor r, g, b
        Call pb.DrawLine(pb.PageWidth - x - 1.5 * size, y, pb.PageWidth - x - 1.5 * size, y + 1)
        page = page + 1
        
        If maxpage Mod 2 = 0 Then       '��ҳ���Ƿ�Ϊż��
            Do While page <= maxpage
                Call pb.NewPage
                pb.SetPageSize "A4"
                pb.SetTextSize size
                pb.SetTextColor r, g, b
                If page Mod 2 <> 0 Then
                    Call pb.DrawLine(pb.PageWidth - x - 1.5 * size, y, pb.PageWidth - x - 1.5 * size, y + 1)
                Else
                    Call pb.DrawLine(x, y, x, y + 1)
                End If
                page = page + 1
            Loop
        Else                               '��ż�������һҳ��Ҫ��ǰ��һҳ�հ�ҳ
            Do While page <= maxpage
                If page <> maxpage Then
                    Call pb.NewPage
                    pb.SetPageSize "A4"
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    If page Mod 2 <> 0 Then
                        Call pb.DrawLine(pb.PageWidth - x - 1.5 * size, y, pb.PageWidth - x - 1.5 * size, y + 1)
                    Else
                        Call pb.DrawLine(x, y, x, y + 1)
                    End If
                Else                        '���һҳ
                    Call pb.NewPage         '�հ�ҳ
                    pb.SetPageSize "A4"
                    Call pb.NewPage
                    pb.SetPageSize "A4"
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    Call pb.DrawRotatedText(x, y, x, y + 1)
                End If
                page = page + 1
            Loop
        End If
        
        result = pb.SaveToFile(Text4.Text)
        If result = 0 Then MsgBox ("����ʧ�ܣ����ȹر��ļ�")
    End If
    SB1.Panels(1).Text = "��� ��ʱ��" & Timer - time & "s"
    Timer1.Enabled = True
End Sub

Private Sub Command6_Click()
    cd1.ShowOpen
    Text4.Text = cd1.FileName
End Sub

Private Sub Command7_Click()
    SB1.Panels(1).Text = "������..."
    If check2() = True Then
        time = Timer
        Dim pb As New DebenuPDFLibraryAX1016.PDFLibrary
        Dim size As Double
        Dim page As Integer
        Dim maxpage As Integer
        Dim r%, g%, b%
        Dim x%, y%
        
        r = (cd1.Color And &HFF) Mod 256
        g = ((cd1.Color And &HFF00) \ &H100) Mod 256
        b = ((cd1.Color And &HFF0000) \ &H10000) Mod 256
        Call pb.UnlockKey("j87ig3k84fb9eq9dy34z7u66y")
        
        size = Val(Text3.Text)
        angle = Val(Text8.Text)
        pb.SetOrigin 1
        maxpage = Val(Text5.Text)
        page = 1
    
        pb.SelectPage page     '��һҳ
        pb.SetPageSize "A4"
        x = pb.PageWidth * Slider1.Value / 100
        y = pb.PageHeight * Slider2.Value / 100
        pb.SetTextSize size
        pb.SetTextColor r, g, b
        Call pb.DrawRotatedText(pb.PageWidth - x - 1.5 * size, y, angle, String(4 - Len(Str(page + 1)), "0") & page + 1)
        page = page + 1
        
        If maxpage Mod 2 = 0 Then       '��ҳ���Ƿ�Ϊż��
            Do While page <= maxpage
                Call pb.NewPage
                pb.SetPageSize "A4"
                pb.SetTextSize size
                pb.SetTextColor r, g, b
                If page Mod 2 <> 0 Then
                    Call pb.DrawRotatedText(pb.PageWidth - x - 1.5 * size, y, angle, String(4 - Len(Str(page + 1)), "0") & page + 1)
                Else
                    Call pb.DrawRotatedText(x, y, angle, String(4 - Len(Str(page - 1)), "0") & page - 1)
                End If
                page = page + 1
            Loop
        Else                               '��ż�������һҳ��Ҫ��ǰ��һҳ�հ�ҳ
            Do While page <= maxpage
                If page <> maxpage Then
                    Call pb.NewPage
                    pb.SetPageSize "A4"
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    If page Mod 2 <> 0 Then
                        Call pb.DrawRotatedText(pb.PageWidth - x - 1.5 * size, y, angle, String(4 - Len(Str(page + 1)), "0") & page + 1)
                    Else
                        Call pb.DrawRotatedText(x, y, angle, String(4 - Len(Str(page - 1)), "0") & page - 1)
                    End If
                Else                        '���һҳ
                    Call pb.NewPage         '�հ�ҳ
                    pb.SetPageSize "A4"
                    Call pb.NewPage
                    pb.SetPageSize "A4"
                    pb.SetTextSize size
                    pb.SetTextColor r, g, b
                    Call pb.DrawRotatedText(x, y, angle, String(4 - Len(Str(page)), "0") & page)
                End If
                page = page + 1
            Loop
        End If
        
        result = pb.SaveToFile(Text4.Text)
        If result = 0 Then MsgBox ("����ʧ�ܣ����ȹر��ļ�")
    End If
    SB1.Panels(1).Text = "��� ��ʱ��" & Timer - time & "s"
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    App.Title = ""
    cd1.Color = 0
    Box1.Scale (100, 100)-(0, 0)
    Color1.BackColor = cd1.Color
    Text7 = "����ģ����һ��ֻ����ҳ���pdf�ļ����������������Ѿ���ӡ�õ��ļ����б��롣�������ļ�װ���ӡ��ֽ����ӡ��ģ��"
    '����ʱ�Զ�ע��ؼ�
    Dim sysdir$, dirlen%
    sysdir = Space(50)
    dirlen = GetSystemDirectory(sysdir, 50)
    sysdir = Left(sysdir, dirlen)
    tdir = Dir(sysdir & "\pdf2parts.dll")
    If tdir = "" Then
        On Error GoTo ERRmsg
        Call FileCopy(App.Path & "\pdf2parts.dll", sysdir & "\pdf2parts.dll")
        Shell App.Path & "\regsvr32.exe /s " & sysdir & "\pdf2parts.dll"
        MsgBox "ע��ؼ��ɹ���"
    End If
    Exit Sub
ERRmsg:
    MsgBox "�����޷�ע��ؼ�" & vbCrLf & vbCrLf & "�״��������Թ���Ա������б�����"
    End
End Sub

Sub showpos()
    Box1.Cls
    Dim pos As Integer
    Dim fin As Integer
    pos = 100 - Slider2.Value
    fin = 100 - Slider2.Value + 2
    Do While pos <= fin
        Box1.Line (100 - Slider1.Value, pos)-(100 - Slider1.Value + 6, pos), cd1.Color
        pos = pos + 1
    Loop
End Sub

Function check1() As Boolean
    If Text1 = "" Then
        MsgBox "������PDF�ļ�"
        check1 = False
    ElseIf Text6 = "" Then
        MsgBox "��������ʼ����"
        check1 = False
    ElseIf Text3 = "" Then
        MsgBox "����������Ĵ�С"
        check1 = False
    Else
        check1 = True
    End If
End Function

Function check2() As Boolean
    If Text4 = "" Then
        MsgBox "��������Ч�ı���·��"
        check2 = False
    ElseIf Text5 = "" Then
        MsgBox "���������ҳ��"
        check2 = False
    ElseIf Text5 > 999 Then
        MsgBox "ҳ�����ܳ���999"
        check2 = False
    ElseIf Text3 = "" Then
        MsgBox "����������Ĵ�С"
        check2 = False
    Else
        check2 = True
    End If
End Function

Private Sub Slider1_Change()
    Call showpos
End Sub

Private Sub Slider2_Change()
    Call showpos
End Sub


Private Sub Tab1_Click()
    MsgBox "         PdfPager v1.2  2018.9" & vbCrLf & vbCrLf & "Pdf�Զ�������" _
    , vbOKOnly, "����"
End Sub

Private Sub Timer1_Timer()
    SB1.Panels(1).Text = "�ȴ���"
    Timer1.Enabled = False
End Sub
