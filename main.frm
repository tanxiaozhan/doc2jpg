VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DOC2JPG 
   Caption         =   "DOC2JPG"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7665
   StartUpPosition =   2  '��Ļ����
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   165
      TabIndex        =   6
      Top             =   465
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�������ļ���"
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   2325
      Value           =   1  'Checked
      Width           =   1530
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   6930
      Top             =   2805
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.doc"
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   5535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��       ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4095
      TabIndex        =   1
      Top             =   3090
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ת        ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   975
      TabIndex        =   0
      Top             =   3075
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "��ѡ��Word�ļ����ڵ��ļ��У�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1425
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   720
      Picture         =   "main.frx":08CA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "DOC2JPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ----====   API   Declarations   ====----
'Globally Unique Identifier������
Private Type BrowseInfo
          hWndOwner   As Long
          pIDLRoot   As Long
          pszDisplayName   As Long
          lpszTitle   As Long
          ulFlags   As Long
          lpfnCallback   As Long
          lParam   As Long
          iImage   As Long
End Type


Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
'Picture Descriptor������ MSDN(Eng)����
Private Type PICTDESC
    cbSizeofstruct As Long
    picType As Long
    hbitmap As Long
    hpal As Long
    unused_wmf_yExt As Long
End Type
Private Type rect
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
Private Type METAFILEPICT
        mm As Long
        xExt As Long
        yExt As Long
        Hmf As Long
End Type
Private Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type SIZEL
    cx As Long
    cy As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long              '�إå��`�Υ�����
    biWidth As Long             '��(�ԥ�����gλ)
    biHeight As Long            '�ߤ�(�ԥ�����gλ)
    biPlanes As Integer         '���ˣ�
    biBitCount As Integer       '1�ԥ����뤢����Υ���`�ӥå���
    biCompression As Long       '�R�s����
    biSizeImage As Long         '�ԥ�����ǩ`����ȫ�Х�����
    biXPelsPerMeter As Long     '0�ޤ���ˮƽ�����
    biYPelsPerMeter As Long     '0�ޤ��ϴ�ֱ�����
    biClrUsed As Long           'ͨ����0
    biClrImportant As Long      'ͨ����0
End Type
Private Type RGBQUAD
    rgbBlue As Byte             '��Ν⤵
    rgbGreen As Byte            '�v�Ν⤵
    rgbRed As Byte              '��Ν⤵
    rgbReserved As Byte         'δʹ��(����0)
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type ENHMETAHEADER
        itype As Long
        nSize As Long
        rclBounds As RECTL
        rclFrame As RECTL '0.01mm�gλ
        dSignature As Long
        nVersion As Long
        nBytes As Long
        nRecords As Long
        nHandles As Integer
        sReserved As Integer
        nDescription As Long
        offDescription As Long
        nPalEntries As Long
        szlDevice As SIZEL
        szlMillimeters As SIZEL
End Type

Private Type GdiplusStartupInput
    GdiplusVersion   As Long
    DebugEventCallback   As Long
    SuppressBackgroundThread   As Long
    SuppressExternalCodecs   As Long
End Type

Private Type EncoderParameter
    Guid   As Guid
    NumberOfValues   As Long
        type   As Long
    value   As Long
End Type

Private Type EncoderParameters
    count   As Long
    Parameter   As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
        token As Long, _
        inputbuf As GdiplusStartupInput, _
        Optional ByVal outputbuf As Long = 0) As Long

Private Declare Function GdiplusShutdown Lib "gdiplus" ( _
        ByVal token As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" ( _
        ByVal hbm As Long, _
        ByVal hpal As Long, _
        Bitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
        ByVal Image As Long) As Long

Private Declare Function GdipSaveImageToFile Lib "gdiplus" ( _
        ByVal Image As Long, _
        ByVal FileName As Long, _
        clsidEncoder As Guid, _
        encoderParams As Any) As Long

Private Declare Function CLSIDFromString Lib "ole32" ( _
        ByVal str As Long, _
        id As Guid) As Long
        
        
'====== ���� ======
Private Const PICTYPE_BITMAP = 1        'pictdesc���뤨��picture�Υ�����
Private Const DIB_RGB_COLORS = 0&

Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As Long
Private Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal Hmf As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As METAFILEPICT) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hEmf As Long, lpRect As rect) As Long
Private Declare Function SetWinMetaFileBitsByNull Lib "gdi32" Alias "SetWinMetaFileBits" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Private Declare Function GetEnhMetaFileHeader Lib "gdi32" ( _
  ByVal hEmf As Long, _
  ByVal MetaHeaderSize As Long, _
  ByRef METAHEADER As ENHMETAHEADER) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" _
    (ByVal hdc As Long, pbmi As BITMAPINFO, ByVal iUsage As Long, _
    ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, _
                                                   ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, _
                                                   lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
        (ByVal hdc As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" _
        (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" _
        (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
        (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" _
        (ByVal hwnd As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" _
        (lpPictDesc As PICTDESC, riid As Guid, _
        ByVal fOwn As Long, lplpvObj As Object) As Long
Private Declare Function SelectObject Lib "gdi32" _
        (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
Private Declare Function ReleaseDC Lib "user32" _
        (ByVal hwnd As Long, ByVal hdc As Long) As Long

Const PICTYPE_ENHMETAFILE = 4

Private Declare Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
Const CF_ENHMETAFILE = 14
Private Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (LpBrowseInfo As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)


Private Sub clipEmf2bmp(ByVal pageNo As Integer)
  Dim hbmp As Long
  Dim hbmpOld As Long
  Dim hdc As Long, hdcDesktop As Long
  Dim hEmf As Long '�����᥿�ե�����Υϥ�ɥ�
  Dim r As rect '�軭�����I��
  Dim strFileName As String
  Dim mh As ENHMETAHEADER 'ȡ�ýY���Υ᥿�ե�����إå�
  Dim emfWidth As Long, emfHeight As Long
  Dim bmpInfo As BITMAPINFO
  Dim pic As StdPicture 'Picture�ץ�ѥƥ��Υǩ`����
 
  'Selection.Copy
  
  If OpenClipboard(0) Then
    hEmf = GetClipboardData(CF_ENHMETAFILE)
    ' �ϥ�ɥ���}�u���Ƥ���ʹ�ä���
    
    hEmf = CopyEnhMetaFile(hEmf, vbNullString)
    CloseClipboard
  End If
  If hEmf = 0 Then
    MsgBox "emfȡ��ʧ��"
    Exit Sub  ' ʧ��
  End If
   '�إå���ȡ��
  GetEnhMetaFileHeader hEmf, Len(mh), mh
  'Excel�Ǉ�Β�����N�긶�����Ԫ������sС����롣
    '�������Ԫ����ν����dpi�O����ӳ������餷��
  '72dpi�λ���ΤȤ����ݥ����=1/72�Ⱥ��¤��뤿�ᡢӋ���ϤΥ������ȡ���������·��
  '�ץ�ѥƥ��Ǳ�ʾ�����編��cm�gλ�������¤��롣ָ�����o�����96dpi��Ҋ�ʤ���롣
  With mh
     emfWidth = .rclBounds.Right - .rclBounds.Left
     emfHeight = .rclBounds.Bottom - .rclBounds.Top
  End With
  hdcDesktop = GetDC(0)
  hdc = CreateCompatibleDC(hdcDesktop)
  
'    hbmp = CreateCompatibleBitmap(hdc, emfWidth, emfHeight) ����Ǥϰ��\����
'    http://hpcgi1.nifty.com/MADIA/VBBBS/wwwlng.cgi?print+200504/05040072.txt
  With bmpInfo.bmiHeader '��������ڻ�
    .biSize = 40
    .biWidth = emfWidth
    .biHeight = emfHeight
    .biPlanes = 1
    .biBitCount = 24 '�����ӥå�
    .biCompression = 0 'BI_RGB
    .biSizeImage = 0 'BI_RGB�Εr�ϣ�
    .biClrUsed = 0
  End With
  bmpInfo.bmiColors.rgbBlue = 202
  bmpInfo.bmiColors.rgbGreen = 200
  bmpInfo.bmiColors.rgbRed = 100
  
  Dim i, pdat As Long
  
  hbmp = CreateDIBSection(hdc, bmpInfo, DIB_RGB_COLORS, _
            pdat, 0, 0) 'DIB����
  hbmpOld = SelectObject(hdc, hbmp)
  '�軭�I����O��
  r.Left = 0
  r.Top = 0
  r.Right = emfWidth
  r.Bottom = emfHeight
    
    ' Create the device context.
    'mem_dc = CreateCompatibleDC(hdc)

    ' Create the bitmap.
    'mem_bm = CreateCompatibleBitmap(mem_dc, wid, hgt)

    ' Make the device context use the bitmap.
    'orig_bm = SelectObject(mem_dc, mem_bm)

    ' Give the device context a white background.
    SelectObject hdc, GetStockObject(0) 'WHITE_BRUSH
    Rectangle hdc, 0, 0, emfWidth, emfHeight
    SelectObject hdc, GetStockObject(5)

    ' Draw the on the device context.
    'SelectObject hdc, GetStockObject(7)
    'MoveToEx hdc, 0, 0, ByVal 0&
    'LineTo hdc, 100, 100
    'MoveToEx hdc, 0, 100, ByVal 0&
    'LineTo hdc, 100, 0
  
  
  
  
  '�����᥿�ե�������軭
  
  Call PlayEnhMetaFile(hdc, hEmf, r)
  
  Set pic = GetPictureObject(hbmp)
  SaveJPG pic, "e:\pic" & pageNo & ".jpg"
  'SavePicture pic, "e:\Test.bmp"
  SelectObject hdc, hbmpOld
  DeleteObject hbmp
  DeleteDC hdc
  DeleteEnhMetaFile hEmf ' ��Ҫ������
End Sub


'====================================================
' HBITMAP����Picture���֥������Ȥ����ɤ����v��
'������BitMap�Υϥ�ɥ�
Private Function GetPictureObject(ByVal hbmp As Long) As Object
 
    Dim iid As Guid     'Globally Unique Identifier�ͤΉ���iid
    Dim pd As PICTDESC  'Picture Descriptor�������ͤΉ���pd
    '�ӥåȥޥåפΥϥ�ɥ뤬0�ʤ顢�K��
    If hbmp = 0 Then Exit Function
    'GUID�͘�����iid�Υ��Ф��O��
    With iid
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    'Picture Descriptor��������O��
    With pd
        .cbSizeofstruct = Len(pd)   'PICTDESC structure�Υ�����
        .picType = PICTYPE_BITMAP   'picture�Υ����ף�PICTYPE�В����꣩
        .hbitmap = hbmp             '�ӥåȥޥåפΥϥ�ɥ�
    End With
    'PICDESC��������O����������Ԫ�˥ԥ�����`���֥������Ȥ����ɡ�
    'OleCreatePictureIndirect(udtPICTDESC, udtGUID, True, NewPic)
    OleCreatePictureIndirect pd, iid, 1, GetPictureObject
 
End Function

'���ޤ���emf�ե������bmp�ˉ�Q����
Private Sub emf2bmp()
  Dim hbmp As Long
  Dim hbmpOld As Long
  Dim hdc As Long, hdcDesktop As Long
  Dim hEmf As Long '�����᥿�ե�����Υϥ�ɥ�
  Dim r As rect '�軭�����I��
  Dim strFileName As String
  Dim mh As ENHMETAHEADER 'ȡ�ýY���Υ᥿�ե�����إå�
  Dim emfWidth As Long, emfHeight As Long
  Dim bmpInfo As BITMAPINFO
  Dim pic As StdPicture 'Picture�ץ�ѥƥ��Υǩ`����
  
  strFileName = "c:\saveEmfTest.emf"
    '�����᥿�ե�����Υ��`�ץ�
   hEmf = GetEnhMetaFile(strFileName)
   '�إå���ȡ��
  GetEnhMetaFileHeader hEmf, Len(mh), mh
  With mh
     '�gλ��pixcel
     emfWidth = .rclBounds.Right - .rclBounds.Left
     emfHeight = .rclBounds.Bottom - .rclBounds.Top
     '.rclFrame.Right - .rclFrame.Left������Υץ�ѥƥ��ǥ������Ȥ��Ʊ�ʾ�����編�Ǥ���
     '��ӛӋ��νY���ϡ�Փ��������һ�w�ˤʤ�
'      emfWidth = (.rclFrame.Right - .rclFrame.Left) * (96 / 25.4) / 100
'      emfHeight = (.rclFrame.Bottom - .rclFrame.Top) * (96 / 25.4) / 100
   End With
   hdcDesktop = GetDC(0)
   hdc = CreateCompatibleDC(hdcDesktop)
'    hbmp = CreateCompatibleBitmap(hdc, emfWidth, emfHeight) ����Ǥϰ��\����
'    http://hpcgi1.nifty.com/MADIA/VBBBS/wwwlng.cgi?print+200504/05040072.txt
  With bmpInfo.bmiHeader '��������ڻ�
    .biSize = 40
    .biWidth = emfWidth
    .biHeight = emfHeight
    .biPlanes = 1
    .biBitCount = 24 '�����ӥå�
    .biCompression = 0 'BI_RGB
    .biSizeImage = 0 'BI_RGB�Εr�ϣ�
    .biClrUsed = 0
  End With
  Dim hDIB As Long
  hbmp = CreateDIBSection(hdc, bmpInfo, DIB_RGB_COLORS, _
            0, 0, 0) 'DIB����
  hbmpOld = SelectObject(hdc, hbmp)
  '�軭�I����O��
  r.Left = 0
  r.Top = 0
  r.Right = emfWidth
  r.Bottom = emfHeight
  
  '�����᥿�ե�������軭
  Call PlayEnhMetaFile(hdc, hEmf, r)
  Set pic = GetPictureObject(hbmp)
  SavePicture pic, "C:\Test.bmp"
  SelectObject hdc, hbmpOld
  DeleteDC hdc
  DeleteEnhMetaFile hEmf
End Sub
        

'   ----====   SaveJPG   ====----

  Public Sub SaveJPG( _
        ByVal pict As StdPicture, _
        ByVal FileName As String, _
        Optional ByVal Quality As Byte = 80)
    Dim tSI     As GdiplusStartupInput
    Dim lRes     As Long
    Dim lGDIP     As Long
    Dim lBitmap     As Long

    '   Initialize   GDI+
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI)

    If lRes = 0 Then

        '   Create   the   GDI+   bitmap
        '   from   the   image   handle
        lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
        If lRes = 0 Then
            Dim tJpgEncoder     As Guid
            Dim tParams     As EncoderParameters

            '   Initialize   the   encoder   GUID
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), _
                    tJpgEncoder

            '   Initialize   the   encoder   parameters
            tParams.count = 1
            With tParams.Parameter                         '   Quality
                '   Set   the   Quality   GUID
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB3505E7EB}"), .Guid
                .NumberOfValues = 1
                .type = 1
                .value = VarPtr(Quality)
            End With

            '   Save   the   image
            lRes = GdipSaveImageToFile( _
                    lBitmap, _
                    StrPtr(FileName), _
                    tJpgEncoder, _
                    tParams)

            '   Destroy   the   bitmap
            GdipDisposeImage lBitmap

        End If

        '   Shutdown   GDI+
        GdiplusShutdown lGDIP

    End If

    If lRes Then
        Err.Raise 5, , "Cannot   save   the   image.   GDI+   Error:" & lRes
    End If

End Sub

Private Sub cmdOpen_Click()
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    txtPath.Text = ""
    
    With udtBI
          '�����������
          .hWndOwner = Me.hwnd
          '����ѡ�е�Ŀ¼
          .ulFlags = BIF_RETURNONLYFSDIRS
    End With
        
        '�����������
        lpIDList = SHBrowseForFolder(udtBI)
        If lpIDList Then
                sPath = String$(256, 0)
                '��ȡ·��
                SHGetPathFromIDList lpIDList, sPath
                '�ͷ��ڴ�
                CoTaskMemFree lpIDList
                iNull = InStr(sPath, vbNullChar)
                If iNull Then
                        sPath = Left$(sPath, iNull - 1)
                        txtPath.Text = sPath
                End If
        End If
End Sub

Private Sub Command1_Click()
    Test
End Sub

Sub Test()
    On Error GoTo err_number
    
    Dim oWord As New Word.Application
    Dim oDoc As New Word.Document
    Dim i As Integer
    
    
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer, NumDirs As Integer
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    
    
    
    
    If txtPath.Text = "" Then
        MsgBox "��ѡ��Word�ļ����ڵ��ļ��С�", , "DOC2JPG"
        Exit Sub
    End If
    
    
        If txtPath.Text <> "" Then
                Screen.MousePointer = vbHourglass
                SearchPath = txtPath.Text       'ѡ�е�Ŀ¼Ϊ��������ʼ·��
                FindStr = "*.doc"               '����doc���͵��ļ�(�˴�����������)
                List1.Clear
                FindFilesAPI SearchPath, FindStr, NumFiles, NumDirs
                Screen.MousePointer = vbDefault
                lblMessage1.Caption = "Ŀ���ļ���   " & List1.ListCount
        End If
    
    
    
    
    

    Clipboard.Clear
    Set oWord = CreateObject("word.application")
    Set oDoc = oWord.Documents.Open("e:\a.doc")
    With oDoc.Application.Selection
        PageCount = .Information(wdNumberOfPagesInDocument)
        For i = 1 To PageCount
        currentpagestart = .GoTo(what:=wdGoToPage, which:=wdGoToNext, Name:=i).Start
        If i = PageCount Then
            currentpageend = oDoc.Content.End
        Else
            currentpageend = .GoTo(what:=wdGoToPage, which:=wdGoToNext, Name:=i + 1).Start
        End If
        oDoc.Range(currentpagestart, currentpageend).Select
        
        .CopyAsPicture
        clipEmf2bmp i
        
   
        Clipboard.Clear
        Next
    End With
err_number: If Err.Number <> 0 Then MsgBox Err.Description
    oDoc.Close False
    oWord.Application.Quit
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
