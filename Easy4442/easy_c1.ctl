VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl Easy4442 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   795
   FillStyle       =   0  'Solid
   Picture         =   "easy_c1.ctx":0000
   ScaleHeight     =   1125
   ScaleMode       =   0  'User
   ScaleWidth      =   803.548
   ToolboxBitmap   =   "easy_c1.ctx":00D8
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
   End
End
Attribute VB_Name = "Easy4442"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim data_reg As Integer         'Buffer per shift-register
Dim bit_cnt As Integer          'Contatore scorrimento bits shift-register
Dim bit_cnt_1 As Integer        'Contatore scorrimento bits PC <-- Card
Dim bit_cnt_2 As Integer        'Contatore scorrimento bits PC --> Card
Dim atr_cnt As Integer          'Puntatore ATR
Dim s_m_cnt As Integer          'Puntatore secrity memory
Dim p_m_cnt As Integer          'Puntatore protection memory
Dim PIN_cnt As Integer          'Puntatore PIN
Dim Addr_Cnt As Integer         'Puntatote indirizzo memoria carta
Dim Pot_2 As Integer            'Registro temporaneo
Dim card_RST As Integer         'Pin RESET (0--> Normale funz. 1--> Reset carta)
Dim card_CLK As Integer         'Pin CLOCK della carta
Dim card_I_O As Integer         'PIN I/O della carta
Dim RXreg As Integer            'Dato ricevuto dalla carta (Buffer ric.)
Dim TXreg As Integer            'Dato inviato alla carta   (Buffer trasm.)
Dim Card_ADDR As Integer        'Indirizzo locazione carta
Dim card_CMND As Integer        'Comando da inviare alla carta
Dim Card_DATA As Integer        'Dato da inviare alla carta
Dim Processing_Stat As Boolean  'Se "True" indica che è in corso un "processing"
Dim Outgoing_mode As Boolean    'Specifica se il comando prevede l'invio di dati
Dim Singolo_Byte As Boolean     'Specifica se leggere un singolo byte o tutta la memoria
Dim card_present As Boolean     'Indica se la carta è inserita nel lettore
Dim c_pin_ok As Boolean         'Indica se il controllo del PIN è avvenuto con successo
Dim Error_Counter As Integer    'Error Counter
Dim Numero_Tentativi As Integer 'Numero di tentativi di verifica PIN disponibili(0..3)
Dim card_memory(255) As Integer 'Main memory (SLE4442 256 bytes; 0-->31 Security memory)
Dim ATR(3) As Integer           'Answer to reset (4 bytes)
Dim Sec_mem(3) As Integer       'Security memory (SLE4442 4 bytes)
Dim Prot_Mem(3) As Integer      'Protection memory (SLE4442 4 bytes)
Dim PIN(2) As Variant           'Codice PIN


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
'Scrivi_Protection_Memory --> Scrive il contenuto di card_DATA all'indirizzo card_ADDR
'                             e, se il valore presente presente in card_DATA è uguale
'                             al valore attuale del registro indirizzato da card_ADDR,
'                             tale registro non potrà più essere modificato.
'
'Verifica_PIN --> Esegue la comparazione tra i registri PIN(1..3) e i relativi
'                   registri EEPROM interni alla carta. Ad ogni verifica viene decrementato
'                   il numero dei tentativi (numero_tentativi) di una unità (3 totali).
'                   Se la comparazione è avvenuta con successo è possibile scrivere nella
'                   main memory della carta a tutti gli indirizzi (tranne i primi 0..31
'                   se protetti) ed il numero dei tentativi viene ripristinato al valore
'                   iniziale (3).
'
'
'
'
'                        ***** Elenco delle proprietà del modulo *****
'
'
'Main_Memory --> (Sola lettura) Restituisce il contenuto della Main Memory sotto forma
'               di matrice Main_Memory(Memory_ADDR) se la proprietà Single_Byte=True,
'               mentre restituisce una matrice Main_Memory (Memory_ADDR ...255) se
'               la proprietà Single_Byte=False.
'
'Security_Memory --> (sola lettura) Restituisce il contenuto della Security Memory
'
'Protection_Memory --> (sola lettura) Restituisce il contenuto della Protection Memory
'
'Memory_ADDR --> (Lettura - Scrittura) Restituisce o imposta l'indirizzo di una locazione
'                di memoria (Main_Memory, Security_Memory, Protection_Memory)
'
'Memory_DATA --> (Lettura - Scrittura) Restituisce o imposta il dato passato alla carta.
'
'Init_Com_Port --> (sola scrittura) Imposta il numero della porta seriale alla quale
'                   è collegato l'EASY CHECK e configura i parametri di comunicazione.
'
'Close_Com_Port --> (sola chiamata) Chiude la porta seriale precedentemente
'                   specificata nella proprietà Init_Com_Port.
'
'PIN_OK_flag --> (sola lettura) E' un flag (True o False) che se True indica che la
'                   verifica del codice PIN è avvenuta con successo.
'
'CARD_INSERT_flag --> (sola lettura) E' un flag (True o False) che se True indica
'                     che la carta è inserita nel lettore.
'
'Errors --> (sola lettura) Indica le rimanenti possibilità di errore nella
'           verifica del codice PIN.
'
'Answer_To_Reset --> (sola lettura) Restituisce i primi 4 bytes della Main Memory corrispondenti
'                   ai byte dell' Answer to reset.
'
'Card_PIN --> (lettura e scrittura) E' la matrice Card_PIN(1..3) che contiene il codice PIN.
'               Il byte Card_PIN(0) è il registro error counter della carta.
'
'
'                   ***** Utilizzo del modulo Easy4442 *****
'
'       1) Lanciare la routine Easy4442n.Init_Com_Port=[com]
'       1a)Eseguire un poolling per verificare la presenza della carta nel lettore
'           controllando il falg Easy4442n.CARD_INSERT_flag (se True la carta è inserita)
'       2) Lanciare la routine Easy4442n.Invia_ATR per resettare la carta e leggere i bytes dell'ATR con
'               valore_ATR(0..3) = Easy4442n.Answer_To_Reset(0..3)
'
'       3) Eseguire una lettura, una scrittura o verificare il codice PIN:
'
'           a)Eseguire una lettura di un byte della Main Memory
'               Easy4442n.Memory_ADDR = [indirizzo cella 0..255]
'               Easy4442n.Single_Byte = True
'               Easy4442n.Leggi_Main_Memory
'               valore_cella = Easy4442n.Main_Memory([indirizzo cella])
'
'           b)Eseguire la lettura della Main Memory da un indirizzo all'ultima locazione
'               Easy4442n.Memory_ADDR = [indirizzo di partenza 0..255]
'               Easy4442n.Single_Byte = False
'               Easy4442n.Leggi_Main_Memory
'               valore_cella[indirizzo di partenza..255] = Easy4442n.Main_Memory([indirizzo di partenza..255])
'
'           c)Eseguire la lettura della Security Memory
'               Easy4442n.Leggi_Security_Memory
'               error_counter = Easy4442n.Security_Memory(0)
'               PIN1 = Easy4442n.Security_Memory(1)
'               PIN2 = Easy4442n.Security_Memory(2)
'               PIN3 = Easy4442n.Security_Memory(3)
'
'           d)Eseguire la lettura della Protection Memory
'               Easy4442n.Leggi_Prtection_Memory
'               valore_protecton_memory(0..3) = Easy4442n.Protection_Memory(0..3)
'
'           e)Eseguire la scrittura in una cella della Main Memory (Richiede precedente verifica PIN)
'               Easy4442n.Memory_ADDR = [indirizzo della cella]
'               Easy4442n.Memory_DATA = [dato da scrivere nella cella]
'               Easy4442n.Scrivi_Main_Memory
'
'           f)Eseguire la verifica del codice PIN
'               Easy4442n.Card_PIN(0) = [primo numero PIN 0..255]
'               Easy4442n.Card_PIN(1) = [secondo numero PIN 0..255]
'               Easy4442n.Card_PIN(2) = [terzo numero PIN 0..255]
'               Easy4442n.Verifica_PIN
'
'               Se il PIN immesso è corretto, il flag Easy4442n.PIN_OK_flag sarà=True,
'               altrimenti risulterà uguale a False.
'
'           g)Proteggere un Byte(0..31) della Main Memory (Richiede precedente verifica PIN)
'               Easy4442n.Memory_ADDR = [indirizzo cella da proteggere] (solo da 0..31)
'               Easy4442n.Memory_DATA = [dato presente all'indirizzo specificato]
'               Easy4442n.Scrivi_Protection_Memory
'


Private Sub delay1()
i = 1
Do While i > 0
    i = i - 1
Loop
End Sub
Private Sub crea_data_reg()
data_reg = 0
If card_RST = 1 Then
    data_reg = (data_reg Or 2)
End If
If card_CLK = 1 Then
    data_reg = (data_reg Or 8)
End If
If card_I_O = 1 Then
    data_reg = (data_reg Or 16)
End If
End Sub
Private Sub aggiorna_shift_register()
' ***** Il bit 7 è sempre messo a 1 per poter attivare la linea RCK dopo 8 bit *****
Call crea_data_reg
bit_cnt = 7
MSComm1.RTSEnable = True    ' Bit 7 a 1
Call s_r_clock
ciclo_scritt_s_r:
Pot_2 = (2 ^ (bit_cnt - 1))
If (data_reg And Pot_2) = Pot_2 Then
    MSComm1.RTSEnable = True
        Else: MSComm1.RTSEnable = False
End If
Call s_r_clock
bit_cnt = (bit_cnt - 1)
If bit_cnt > 0 Then
    GoTo ciclo_scritt_s_r
End If
MSComm1.RTSEnable = True ' Mantieni acceso il circuito
End Sub
Private Sub ricevi_BYTE()
RXreg = 0
bit_cnt_1 = 0                'Il primo bit inviato dalla carta è il bit0
ric_b_1:
card_CLK = 1
Call aggiorna_shift_register
If MSComm1.CTSHolding = False Then
    RXreg = (RXreg Or (2 ^ (bit_cnt_1)))
End If
card_CLK = 0
Call aggiorna_shift_register
bit_cnt_1 = (bit_cnt_1 + 1)
If bit_cnt_1 < 8 Then
    GoTo ric_b_1
End If
End Sub
Private Sub trasmetti_BYTE()
bit_cnt_2 = 0                'Il primo bit inviato alla carta è il bit0
tra_b_1:
If (TXreg And (2 ^ (bit_cnt_2))) = (2 ^ (bit_cnt_2)) Then
    card_I_O = 1
     Else: card_I_O = 0
End If
card_CLK = 1
Call aggiorna_shift_register
card_CLK = 0
Call aggiorna_shift_register
bit_cnt_2 = (bit_cnt_2 + 1)
If bit_cnt_2 < 8 Then
    GoTo tra_b_1
End If
End Sub
Private Sub s_r_clock()
MSComm1.DTREnable = True
Call delay1
MSComm1.DTREnable = False
Call delay1
End Sub
Public Sub invia_ATR()
card_RST = 1
Call aggiorna_shift_register
card_CLK = 1
Call aggiorna_shift_register
card_CLK = 0
Call aggiorna_shift_register
card_RST = 0
Call aggiorna_shift_register
Call ricevi_BYTE
ATR(0) = RXreg
Call ricevi_BYTE
ATR(1) = RXreg
Call ricevi_BYTE
ATR(2) = RXreg
Call ricevi_BYTE
ATR(3) = RXreg
End Sub
Private Sub invia_break()
card_RST = 1
Call aggiorna_shift_register
Call aggiorna_shift_register
card_RST = 0
Call aggiorna_shift_register
End Sub

Private Sub invia_comando()
'           ***** Sincronismo *****
card_CLK = 1
Call aggiorna_shift_register
card_I_O = 0                                'START from IFD
Call aggiorna_shift_register
card_CLK = 0
Call aggiorna_shift_register
'           ***** Invia CONTROL *****
TXreg = card_CMND
Call trasmetti_BYTE
'           ***** Invia ADDRESS *****
TXreg = Card_ADDR
Call trasmetti_BYTE
'           ***** Invia DATA *****
TXreg = Card_DATA
Call trasmetti_BYTE
'           ***** Sincronismo *****
card_I_O = 0
card_CLK = 0
Call aggiorna_shift_register
card_CLK = 1
Call aggiorna_shift_register
card_I_O = 1                                'STOP from IFD
Call aggiorna_shift_register
If Outgoing_mode = False Then
    card_I_O = 0                            'START of Processing
    Call aggiorna_shift_register
    Processing_Stat = True
    card_CLK = 0
    Call aggiorna_shift_register
    card_I_O = 1
    Call aggiorna_shift_register
inv_c_1:
        If MSComm1.CTSHolding = True Then
            card_CLK = 1
            Call aggiorna_shift_register
            card_CLK = 0
            Call aggiorna_shift_register
            GoTo inv_c_1
            Processing_Stat = False         'END of Processing
        End If
Else
    card_CLK = 0                            'START of Outgoing Data
    Call aggiorna_shift_register

            If card_CMND = 48 Then
              If Singolo_Byte = True Then
               Call ricevi_BYTE
               Call invia_break
               card_memory(Card_ADDR) = RXreg
              Else:
inv_c_2:
               For Card_ADDR = Card_ADDR To 255
               Call ricevi_BYTE
               card_memory(Card_ADDR) = RXreg
               Addr_Cnt = Card_ADDR
               Next Card_ADDR
              End If                        'END of Outgoing Data
            End If
            If card_CMND = 49 Then
              For s_m_cnt = 0 To 3
              Call ricevi_BYTE
              Sec_mem(s_m_cnt) = RXreg
              Next s_m_cnt
            End If
            If card_CMND = 52 Then
              For p_m_cnt = 0 To 3
              Call ricevi_BYTE
              Prot_Mem(p_m_cnt) = RXreg
              Next p_m_cnt
            End If
End If
End Sub
Public Property Let Init_Com_Port(com_port As Integer)
MSComm1.CommPort = com_port
If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
    MSComm1.DTREnable = False
    MSComm1.RTSEnable = True
    data_reg = 0
    card_RST = 0
    card_CLK = 0
    card_I_O = 1
    Call aggiorna_shift_register
End If
End Property
Public Sub Close_Com_Port()
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
End Sub
Public Property Get PIN_OK_flag() As Boolean
If card_present = False Then
c_pin_ok = False
End If
PIN_OK_flag = c_pin_ok
End Property
Public Property Get CARD_INSERT_flag() As Boolean
If MSComm1.DSRHolding = False Then
card_present = True
Else: card_present = False
End If
CARD_INSERT_flag = card_present
End Property
Public Property Let Single_byte(Single_byte As Boolean)
Singolo_Byte = Single_byte
End Property
Public Property Get Main_Memory() As Variant
Main_Memory = card_memory
End Property
Public Property Get Memory_ADDR() As Integer
Memory_ADDR = Card_ADDR
End Property
Public Property Let Memory_ADDR(Memory_ADDR As Integer)
Card_ADDR = Memory_ADDR
End Property
Public Property Get Memory_DATA() As Integer
Memory_DATA = Card_DATA
End Property
Public Property Let Memory_DATA(Memory_DATA As Integer)
Card_DATA = Memory_DATA
End Property
Public Property Get Security_Memory() As Variant
Security_Memory = Sec_mem
End Property
Public Property Let Security_Memory(Security_Memory As Variant)
Security_Memory = Sec_mem 'RILEVATO ERRORE
End Property
Public Property Get Answer_To_Reset() As Variant
Answer_To_Reset = ATR
End Property
Public Property Get Protection_Memory() As Variant
Protection_Memory = Prot_Mem
End Property
Public Property Get Errors() As Integer
Errors = Numero_Tentativi
End Property
Public Property Get Card_PIN() As Variant
Card_PIN = PIN
End Property
Public Property Let Card_PIN(Card_PIN As Variant)
For PIN_cnt = 0 To 2
PIN(PIN_cnt) = Card_PIN(PIN_cnt)
Next PIN_cnt
End Property

Public Sub Leggi_Main_Memory()
card_CMND = 48
Outgoing_mode = True
Call invia_comando
End Sub

Public Sub Leggi_Security_Memory()
card_CMND = 49
Singolo_Byte = False
Outgoing_mode = True
Call invia_comando
Error_Counter = (Sec_mem(0) And 7)
Numero_Tentativi = 0
If (Error_Counter And 1) = 1 Then
Numero_Tentativi = (Numero_Tentativi + 1)
End If
If (Error_Counter And 2) = 2 Then
Numero_Tentativi = (Numero_Tentativi + 1)
End If
If (Error_Counter And 4) = 4 Then
Numero_Tentativi = (Numero_Tentativi + 1)
End If
End Sub

Public Sub Leggi_Protection_Memory()
card_CMND = 52
Singolo_Byte = False
Outgoing_mode = True
Call invia_comando
End Sub

Public Sub Verifica_PIN()
Call Leggi_Security_Memory
If (Error_Counter And 1) = 1 Then
Error_Counter = (Error_Counter And 6)
GoTo ver_p_1
End If
If (Error_Counter And 2) = 2 Then
Error_Counter = (Error_Counter And 5)
GoTo ver_p_1
End If
If (Error_Counter And 4) = 4 Then
Error_Counter = (error_conter And 3)
GoTo ver_p_1
End If
msg = "Il numero massimo di errori è stato superato!"
MsgBox msg
GoTo end_v_p
ver_p_1:
Card_ADDR = 0
Card_DATA = Error_Counter
Call Scrivi_Security_Memory
For PIN_cnt = 0 To 2
Card_ADDR = (PIN_cnt + 1)
Card_DATA = PIN(PIN_cnt)
card_CMND = 51
Outgoing_mode = False
Call invia_comando
Next PIN_cnt
Card_ADDR = 0
Card_DATA = 255
Call Scrivi_Security_Memory
Call Leggi_Security_Memory
    If (Sec_mem(0) And 7) = 7 Then
    c_pin_ok = True
    Else: c_pin_ok = False
    End If
end_v_p:
End Sub

Public Sub Scrivi_Security_Memory()
card_CMND = 57
Outgoing_mode = False
Call invia_comando
End Sub

Public Sub Scrivi_Main_Memory()
card_CMND = 56
Outgoing_mode = False
Call invia_comando
End Sub

Public Sub Scrivi_Protection_Memory()
card_CMND = 60
Outgoing_mode = False
Call invia_comando
End Sub


