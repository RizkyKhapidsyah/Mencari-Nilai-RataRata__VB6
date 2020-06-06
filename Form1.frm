VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencari Nilai Rata-Rata"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer, Total As Long
'Karena kita menggunakan array statis, maka kita
'harus menggunakan constanta tertentu untuk
'menentukan 'banyaknya elemen array. Jika kita
'menggunakan:

'Dim JlhBil As Integer
'(misalnya), maka akan terjadi error saat run-time
'yang menunjuk baris: Dim arrData(JlhBil) As Integer

 Const JlhBil As Integer = 10
Dim arrData(JlhBil) As Integer
  'Isi elemen array arrData
  arrData(0) = 12
  arrData(1) = 500
  arrData(2) = 92
  arrData(3) = 262
  arrData(4) = 112
  arrData(5) = 152
  arrData(6) = 887
  arrData(7) = 10
  arrData(8) = 120
  arrData(9) = 12
  'Inisialisasi variabel Total
  Total = 0
  'Bersihkan form
  Form1.Cls
  'Iterasi bilangan untuk menjumlah
  For i = 0 To 9
    'Cetak data-nya ke layar
    Print arrData(i)
    'Jumlahkan semua bilangan...
    Total = Total + arrData(i)
  Next i
  'Tampilkan jumlah semua bilangan
  MsgBox "Jumlah seluruh bilangan = " & Total, _
         vbInformation, "Total"
  'Tampilkan nilai rata-rata setelah selesai iterasi
  MsgBox "Rata-rata = " & Total / 10 & "" & vbCrLf & _
         "yang diperoleh dari:" & vbCrLf & _
         "perhitungan: " & Total & "/" & JlhBil & "", _
         vbInformation, "Rata-rata"
End Sub


