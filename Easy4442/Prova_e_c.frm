VERSION 5.00
Object = "{4CBCF0FA-B114-11D4-8D2D-0050BF345302}#1.0#0"; "EASY4442.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Easy4442"
   ClientHeight    =   5265
   ClientLeft      =   3900
   ClientTop       =   990
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Quartz"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Prova_e_c.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7080
   Begin VB.CommandButton Command13 
      BackColor       =   &H000000FF&
      Caption         =   "Proteggi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0080FFFF&
      Caption         =   "Verifica PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   5760
      TabIndex        =   24
      Text            =   "0"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   480
      Index           =   3
      Left            =   5640
      TabIndex        =   23
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   480
      Index           =   2
      Left            =   4800
      TabIndex        =   22
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   480
      Index           =   1
      Left            =   3960
      TabIndex        =   21
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   480
      Index           =   0
      Left            =   3120
      TabIndex        =   20
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FF8080&
      Caption         =   "Leggi p.m."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004040&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Index           =   3
      Left            =   5640
      TabIndex        =   18
      Text            =   "0"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004040&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Index           =   2
      Left            =   4800
      TabIndex        =   17
      Text            =   "0"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004040&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Index           =   1
      Left            =   3960
      TabIndex        =   16
      Text            =   "0"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004040&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Index           =   0
      Left            =   3000
      TabIndex        =   15
      Text            =   "0"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   3
      Left            =   5640
      TabIndex        =   14
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   13
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   12
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FFFF&
      Caption         =   "Leggi s.m."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3375
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF8080&
      Caption         =   "Leggi tutta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000000C0&
      Caption         =   "Scrivi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
      Caption         =   "Leggi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   480
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "ATR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin Progetto2.Easy4442 Easy44421 
      Height          =   495
      Left            =   6360
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3960
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   3240
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      Height          =   975
      Left            =   3840
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   5
      Height          =   4815
      Left            =   120
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "N.tentativi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   27
      Top             =   3240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00004040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      FillColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   3960
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      FillColor       =   &H0000FF00&
      Height          =   495
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim card_addr_l As Integer
Dim L_Pin(2) As Integer
Dim pin_appo As Variant

'                           ****** Elenco routine: ******
'
's_r_clock --> Invia un impulso di clock alla carta
'           (La linea DTR deve essere tenuta sempre bassa e, dopo il clock
'               ritorna a livello basso)
'
'
'crea_data_reg --> Crea il registro che verrà passato allo shift-register(74HC595)
'        Elenco bits:
'                       bit0 --> N.C.
'                       bit1 --> Pin RST della carta
'                       bit2 --> Abilitazione generatore di clock (se bit3 a 0)
'                                interno ed abilitazione linea TXD verso la carta
'                       bit3 --> Se bit2 a 0 è il pin CLK della carta
'                       bit4 --> Se bit2 a 0 è il PIN I/O della carta
'                       bit5 --> N.C.
'                       bit6 --> N.C.
'                       bit7 --> N.C. *Il bit 7 viene ignorato(serve per attivare RCK)
'
'
'invia_ATR --> Invia il comando ATR alla carta e memorizza la risposta in ATR(0..3)
'
'ricevi_BYTE --> Riceve un byte dalla carta e memorizza il contenuto in RXreg
'
'trasmetti_BYTE --> Trasmette un byte (TXreg) alla carta
'
'invia_comando --> Invia un comando alla carta (3 BYTE)
'
'Leggi_Main_Memory --> Legge il contenuto della Main Memory a partire dall'indirizzo
'                   specificato in card_ADDR. Legge un solo byte se il flag "singolo_byte"
'                   è impostato su "True", altrimenti legge tutta la carta da card_ADDR
'                   fino all'ultima locazione (SLE4442 --> 255). I registri letti sono
'                   memorizzati in card_memory(0..255).
'
'
'Scrivi_Main_Memory --> Scrive il contenuto di card_DATA all'indirizzo card_ADDR.
'                       *** Prima di poter scrivere in qualsiasi locazione è
'                       necessario verificare il PIN !!! ***
'
'Leggi_Security_Memory --> Legge il contenuto della security memory (4 bytes) e
'                           memorizza il contenuto in Sec_Mem(0..3).*** I bytes
'                           1,2,3 visualizzano il PIN, mentre il byte 0 è il registro
'                           che contiene l' error counter, dal quale viene ricavato
'                           il registro numero_tentativi.
'
'Scrivi_Security_Memory --> Scrive un byte (card_DATA) nella s.m all' indirizzo card_ADDR.
'
'Leggi_Protection_Memory --> Legge il contenuto della protection memory e memorizza
'                            il contenuto in Prot_Mem(0..3). *** Ogni bit rappresenta
'                            lo stato di protezione del byte relativo della main memory.
'
'Verifica_PIN --> Esegue la comparazione tra i registri PIN(1..3) e i relativi
'                   registri EEPROM interni alla carta. Ad ogni verifica viene decrementato
'                   il numero dei tentativi (numero_tentativi) di una unità (3 totali).
'                   Se la comparazione è avvenuta con successo è possibile scrivere nella
'                   main memory della carta a tutti gli indirizzi (tranne i primi 0..31
'                   se protetti) ed il numero dei tentativi viene ripristinato al valore
'                   iniziale (3).
'


Private Sub Command10_Click()
Call agg_sec_memory
End Sub

Private Sub agg_sec_memory()
Easy44421.Leggi_Security_Memory
For s_m_cnt = 0 To 3
Text5(s_m_cnt).Text = Easy44421.Security_Memory(s_m_cnt)
Next s_m_cnt
Text7.Text = Easy44421.Errors
End Sub

Private Sub Command11_Click()
Easy44421.Leggi_Protection_Memory
For p_m_cnt = 0 To 3
Text6(p_m_cnt).Text = Easy44421.Protection_Memory(p_m_cnt)
Next p_m_cnt
End Sub

Private Sub Command12_Click()
L_Pin(0) = Text5(1).Text
L_Pin(1) = Text5(2).Text
L_Pin(2) = Text5(3).Text
pin_appo = L_Pin()
Easy44421.Card_PIN = pin_appo
Easy44421.Verifica_PIN
Call agg_sec_memory
End Sub

Private Sub Command13_Click()
Easy44421.Memory_ADDR = Text3.Text
Easy44421.Memory_DATA = Text4.Text
Easy44421.Scrivi_Protection_Memory
End Sub

Private Sub Command5_Click()
Easy44421.invia_ATR
For atr_cnt = 0 To 3
Text2(atr_cnt).Text = Easy44421.Answer_To_Reset(atr_cnt)
Next atr_cnt
End Sub

Private Sub Command7_Click()
Easy44421.Memory_ADDR = Text3.Text
Easy44421.Single_byte = True
Easy44421.Leggi_Main_Memory
Text4.Text = Easy44421.Main_Memory(Text3.Text)
End Sub

Private Sub Command8_Click()
Easy44421.Memory_ADDR = Text3.Text
Easy44421.Memory_DATA = Text4.Text
Easy44421.Scrivi_Main_Memory
End Sub

Private Sub Command9_Click()
Text8.Text = ""
Easy44421.Memory_ADDR = Text3.Text
Easy44421.Single_byte = False
Easy44421.Leggi_Main_Memory
For card_addr_l = Text3.Text To 255
  Text8.Text = (Text8.Text & "(" & card_addr_l & ")" & Easy44421.Main_Memory(card_addr_l) & vbCrLf)
Next card_addr_l
End Sub



Private Sub Form_Load()
Easy44421.Init_Com_Port = 2
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
If Easy44421.CARD_INSERT_flag = True Then
    Shape1.FillStyle = 0
    Else: Shape1.FillStyle = 1
End If
If Easy44421.PIN_OK_flag = True Then
    Shape3.FillStyle = 0
    Else: Shape3.FillStyle = 1
End If
End Sub


