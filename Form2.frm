VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Membuat Perhitungan Waktu (Stopwatch)"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form2"
   ScaleHeight     =   3555
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   720
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Star"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalTenthDetik, TotalDetik, TenthDetik, Detik, _
Menit, Jam As Integer
Dim Jam1 As String

Private Sub Command1_Click()
    'Inisialisasi total sepersepuluh detik
    TotalTenthDetik = -1
    'Aktifkan timer
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    'Memulai atau menghentikan timer kembali
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Timer1_Timer()
    'Tambah dengan satu untuk total sepersepuluh detik.
    'Kita mengeset interval Timer menjadi 10, jadi
    'setiap sepersepuluh detik prosedur ini akan
    'dieksekusi
    TotalTenthDetik = TotalTenthDetik + 1
    'Jika TotalTenthSeconds = 10,
    'set kembali menjadi 0.
    TenthDetik = TotalTenthDetik Mod 10
    '10 kali sepersepuluh detik sama dengan 1 detik.
    'int - akan mengembalikan bilangan integer (bulat)
    'dari pecahan 'Contoh: Int(0.9) = 0 menghasilkan 0
    TotalDetik = Int(TotalTenthDetik / 10)
    'Jika variabel Seconds = 60, set kembali menjadi 0
    Detik = TotalDetik Mod 60
    If Len(Detik) = 1 Then
       Detik = "0" & Detik  'Agar selalu dalam dua
                            'digit
    End If
    Menit = Int(TotalDetik / 60) Mod 60
    If Len(Menit) = 1 Then
       Menit = "0" & Menit    'Agar selalu dalam dua
                          'digit
    End If
    Jam = Int(TotalDetik / 3600)
    If Jam < 9 Then
       Jam1 = "0" & Jam       'Agar selalu dalam dua
                      'digit
    End If
    'Tampilkan hasilnya di Label1 (update terus Label1)
    Label1 = Jam1 & ":" & Menit & ":" & Detik & ":" _
             & TenthDetik & ""
End Sub


