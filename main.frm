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
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   165
      TabIndex        =   6
      Top             =   465
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "包含子文件夹"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "退       出"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "转        换"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "请选择Word文件所在的文件夹："
      BeginProperty Font 
         Name            =   "宋体"
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
'Globally Unique Identifier構造体
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
'Picture Descriptor構造体 MSDN(Eng)参照
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
    biSize As Long              'ヘッダーのサイズ
    biWidth As Long             '幅(ピクセル単位)
    biHeight As Long            '高さ(ピクセル単位)
    biPlanes As Integer         '常に１
    biBitCount As Integer       '1ピクセルあたりのカラービット数
    biCompression As Long       '圧縮方法
    biSizeImage As Long         'ピクセルデータの全バイト数
    biXPelsPerMeter As Long     '0または水平解像度
    biYPelsPerMeter As Long     '0または垂直解像度
    biClrUsed As Long           '通常は0
    biClrImportant As Long      '通常は0
End Type
Private Type RGBQUAD
    rgbBlue As Byte             '青の濃さ
    rgbGreen As Byte            '緑の濃さ
    rgbRed As Byte              '赤の濃さ
    rgbReserved As Byte         '未使用(常に0)
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type ENHMETAHEADER
        itype As Long
        nSize As Long
        rclBounds As RECTL
        rclFrame As RECTL '0.01mm単位
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
        
        
'====== 定数 ======
Private Const PICTYPE_BITMAP = 1        'pictdescに与えるpictureのタイプ
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
  Dim hEmf As Long '拡張メタファイルのハンドル
  Dim r As rect '描画する領域
  Dim strFileName As String
  Dim mh As ENHMETAHEADER '取得結果のメタファイルヘッダ
  Dim emfWidth As Long, emfHeight As Long
  Dim bmpInfo As BITMAPINFO
  Dim pic As StdPicture 'Pictureプロパティのデータ型
 
  'Selection.Copy
  
  If OpenClipboard(0) Then
    hEmf = GetClipboardData(CF_ENHMETAFILE)
    ' ハンドルを複製してから使用する
    
    hEmf = CopyEnhMetaFile(hEmf, vbNullString)
    CloseClipboard
  End If
  If hEmf = 0 Then
    MsgBox "emf取得失败"
    Exit Sub  ' 失敗
  End If
   'ヘッダの取得
  GetEnhMetaFileHeader hEmf, Len(mh), mh
  'Excelで図の挿入で貼り付けると元画像より縮小される。
    '→これは元画像の解像度dpi設定を反映さためらしい
  '72dpiの画像のとき、ポイント=1/72と合致するため、計算上のサイズと、エクセルの図の
  'プロパティで表示される寸法（cm単位）が合致する。指定が無ければ96dpiと見なされる。
  With mh
     emfWidth = .rclBounds.Right - .rclBounds.Left
     emfHeight = .rclBounds.Bottom - .rclBounds.Top
  End With
  hdcDesktop = GetDC(0)
  hdc = CreateCompatibleDC(hdcDesktop)
  
'    hbmp = CreateCompatibleBitmap(hdc, emfWidth, emfHeight) これでは白黒画像
'    http://hpcgi1.nifty.com/MADIA/VBBBS/wwwlng.cgi?print+200504/05040072.txt
  With bmpInfo.bmiHeader '構造体初期化
    .biSize = 40
    .biWidth = emfWidth
    .biHeight = emfHeight
    .biPlanes = 1
    .biBitCount = 24 '２４ビット
    .biCompression = 0 'BI_RGB
    .biSizeImage = 0 'BI_RGBの時は０
    .biClrUsed = 0
  End With
  bmpInfo.bmiColors.rgbBlue = 202
  bmpInfo.bmiColors.rgbGreen = 200
  bmpInfo.bmiColors.rgbRed = 100
  
  Dim i, pdat As Long
  
  hbmp = CreateDIBSection(hdc, bmpInfo, DIB_RGB_COLORS, _
            pdat, 0, 0) 'DIB作成
  hbmpOld = SelectObject(hdc, hbmp)
  '描画領域の設定
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
  
  
  
  
  '拡張メタファイルの描画
  
  Call PlayEnhMetaFile(hdc, hEmf, r)
  
  Set pic = GetPictureObject(hbmp)
  SaveJPG pic, "e:\pic" & pageNo & ".jpg"
  'SavePicture pic, "e:\Test.bmp"
  SelectObject hdc, hbmpOld
  DeleteObject hbmp
  DeleteDC hdc
  DeleteEnhMetaFile hEmf ' 必要か不明
End Sub


'====================================================
' HBITMAPからPictureオブジェクトを作成する関数
'引数はBitMapのハンドル
Private Function GetPictureObject(ByVal hbmp As Long) As Object
 
    Dim iid As Guid     'Globally Unique Identifier型の変数iid
    Dim pd As PICTDESC  'Picture Descriptor構造体型の変数pd
    'ビットマップのハンドルが0なら、終了
    If hbmp = 0 Then Exit Function
    'GUID型構造体iidのメンバを設定
    With iid
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    'Picture Descriptor構造体を設定
    With pd
        .cbSizeofstruct = Len(pd)   'PICTDESC structureのサイズ
        .picType = PICTYPE_BITMAP   'pictureのタイプ（PICTYPE列挙体より）
        .hbitmap = hbmp             'ビットマップのハンドル
    End With
    'PICDESC構造体に設定した情報を元にピクチャーオブジェクトを作成。
    'OleCreatePictureIndirect(udtPICTDESC, udtGUID, True, NewPic)
    OleCreatePictureIndirect pd, iid, 1, GetPictureObject
 
End Function

'おまけ　emfファイルを、bmpに変換する
Private Sub emf2bmp()
  Dim hbmp As Long
  Dim hbmpOld As Long
  Dim hdc As Long, hdcDesktop As Long
  Dim hEmf As Long '拡張メタファイルのハンドル
  Dim r As rect '描画する領域
  Dim strFileName As String
  Dim mh As ENHMETAHEADER '取得結果のメタファイルヘッダ
  Dim emfWidth As Long, emfHeight As Long
  Dim bmpInfo As BITMAPINFO
  Dim pic As StdPicture 'Pictureプロパティのデータ型
  
  strFileName = "c:\saveEmfTest.emf"
    '拡張メタファイルのオープン
   hEmf = GetEnhMetaFile(strFileName)
   'ヘッダの取得
  GetEnhMetaFileHeader hEmf, Len(mh), mh
  With mh
     '単位はpixcel
     emfWidth = .rclBounds.Right - .rclBounds.Left
     emfHeight = .rclBounds.Bottom - .rclBounds.Top
     '.rclFrame.Right - .rclFrame.Leftが画像のプロパティでサイズとして表示される寸法である
     '下記計算の結果は、論理サイズと一緒になる
'      emfWidth = (.rclFrame.Right - .rclFrame.Left) * (96 / 25.4) / 100
'      emfHeight = (.rclFrame.Bottom - .rclFrame.Top) * (96 / 25.4) / 100
   End With
   hdcDesktop = GetDC(0)
   hdc = CreateCompatibleDC(hdcDesktop)
'    hbmp = CreateCompatibleBitmap(hdc, emfWidth, emfHeight) これでは白黒画像
'    http://hpcgi1.nifty.com/MADIA/VBBBS/wwwlng.cgi?print+200504/05040072.txt
  With bmpInfo.bmiHeader '構造体初期化
    .biSize = 40
    .biWidth = emfWidth
    .biHeight = emfHeight
    .biPlanes = 1
    .biBitCount = 24 '２４ビット
    .biCompression = 0 'BI_RGB
    .biSizeImage = 0 'BI_RGBの時は０
    .biClrUsed = 0
  End With
  Dim hDIB As Long
  hbmp = CreateDIBSection(hdc, bmpInfo, DIB_RGB_COLORS, _
            0, 0, 0) 'DIB作成
  hbmpOld = SelectObject(hdc, hbmp)
  '描画領域の設定
  r.Left = 0
  r.Top = 0
  r.Right = emfWidth
  r.Bottom = emfHeight
  
  '拡張メタファイルの描画
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
          '设置浏览窗口
          .hWndOwner = Me.hwnd
          '返回选中的目录
          .ulFlags = BIF_RETURNONLYFSDIRS
    End With
        
        '调出浏览窗口
        lpIDList = SHBrowseForFolder(udtBI)
        If lpIDList Then
                sPath = String$(256, 0)
                '获取路径
                SHGetPathFromIDList lpIDList, sPath
                '释放内存
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
        MsgBox "请选择Word文件所在的文件夹。", , "DOC2JPG"
        Exit Sub
    End If
    
    
        If txtPath.Text <> "" Then
                Screen.MousePointer = vbHourglass
                SearchPath = txtPath.Text       '选中的目录为搜索的起始路径
                FindStr = "*.doc"               '搜索doc类型的文件(此处可另作定义)
                List1.Clear
                FindFilesAPI SearchPath, FindStr, NumFiles, NumDirs
                Screen.MousePointer = vbDefault
                lblMessage1.Caption = "目标文件数   " & List1.ListCount
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
