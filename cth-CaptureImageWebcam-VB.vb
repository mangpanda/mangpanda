'@author : mangpanda

'sebelumnya jangan lupa tambahkan/import namespace "System.Runtime.InteropServices" untuk dapat mengakses objek Model objek komponen (COM)
Imports System.Runtime.InteropServices
Public Class Form1
#Region "Constanta"
    Const WM_CAP_START = &H400S
    Const WS_CHILD = &H40000000
    Const WS_VISIBLE = &H10000000


    Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10
    Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11
    Const WM_CAP_EDIT_COPY = WM_CAP_START + 30
    Const WM_CAP_SEQUENCE = WM_CAP_START + 62
    Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23


    Const WM_CAP_SET_SCALE = WM_CAP_START + 53
    Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52
    Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50


    Const SWP_NOMOVE = &H2S
    Const SWP_NOSIZE = 1
    Const SWP_NOZORDER = &H4S
    Const HWND_BOTTOM = 1
#End Region

#Region "Deklarasi Fungsi"
    Declare Function capGetDriverDescriptionA Lib "Avicap32.dll" (ByVal wDriverIndex As Short, ByVal lpszName As String, ByVal cbName As Integer, ByVal lpszVer As String, ByVal cbVer As Integer) As Boolean

    Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Short, ByVal hWnd As Integer, ByVal nID As Integer) As Integer

    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.AsAny)> ByVal lParam As Object) As Integer

    Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Declare Function DestroyWindow Lib "user32" (ByVal hndw As Integer) As Boolean
#End Region

    Dim VideoSource As Integer
    Dim hWND As Integer

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DriverName As String = Space(80) 'variabel untuk menampilkan nama driver webcam kita
        Dim DriverVersion As String = Space(80) 'variabel untuk menampilkan versi driver webcam kita

        'Pada saat aplikasi dijalankan akan otomatis menampilkan driver yang tersedia pada webcam kita yang terpasang
        For i As Integer = 0 To 9
            If capGetDriverDescriptionA(i, DriverName, 80, DriverVersion, 80) Then
                ListBox1.Items.Add(DriverName.Trim)
            End If
        Next
    End Sub

    'Untuk menampilkan disni menggunakan listbox sebagai tool untuk memilih driver dari webcam kita yang selanjutnya akan ditampilkan pada picture box

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        SendMessage(hWND, WM_CAP_DRIVER_DISCONNECT, VideoSource, 0)
        DestroyWindow(hWND)

        VideoSource = ListBox1.SelectedIndex 'mengambil source gambar dari daftar driver yang ada pada listbox

        hWND = capCreateCaptureWindowA(VideoSource, WS_VISIBLE Or WS_CHILD, 0, 0, 0, 0, PictureBox1.Handle.ToInt32, 0)
        If SendMessage(hWND, WM_CAP_DRIVER_CONNECT, VideoSource, 0) Then
            SendMessage(hWND, WM_CAP_SET_SCALE, True, 0)
            SendMessage(hWND, WM_CAP_SET_PREVIEWRATE, 30, 0)
            SendMessage(hWND, WM_CAP_SET_PREVIEW, True, 0)
            'proses dimana capture sorce kita yang diambil dari webcam akan ditampilkan pada picturebox kita
            SetWindowPos(hWND, HWND_BOTTOM, 0, 0, PictureBox1.Width, PictureBox1.Height, SWP_NOMOVE Or SWP_NOZORDER)

        Else
            DestroyWindow(hWND)

        End If
    End Sub
End Class
