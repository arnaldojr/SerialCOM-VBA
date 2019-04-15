VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Software de leitura serial da balaça modelo BK8000II
'Insper - Instituto de pesquisa e esino
'Código adaptado do site Strokescribe.com
'autor: Arnaldo A. Viana Junior
'rev:0.1
'Data: 15/04/2019

Dim x As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim tempo As Integer


Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Private Sub StrokeReader1_CommEvent(ByVal Evt As StrokeReaderLib.Event, ByVal data As Variant)
Select Case Evt
    Case EVT_DISCONNECT
        Debug.Print "Disconnected"
        MsgBox "Desconectado"
    
    Case EVT_CONNECT
        Debug.Print "Connected"
        MsgBox "Conectado"

    Case EVT_DATA
        cell_idx = x
        buf = StrokeReader1.read(Text) 'faz leitura
        Do
            lf = InStr(buf, Chr(10)) 'Most of scanners send ASCII <LF> after each barcode
           
            If lf = 0 Then Exit Do 'Waiting for ASCII <LF> to be received
               
            s = Left(buf, lf - 1) 'Copythe data before <LF> code as the barcode text
            s = Replace(s, Chr(13), "") ' tira o caracter <CR>
            s = Replace(s, Chr(103), "") ' tira o caracter g
            s = Replace(s, Chr(83), "") ' tira o caracter S
            s = Replace(s, Chr(68), "") ' tira o caracter D
            s = Replace(s, "Comando desconhecido!", "") ' ignora a msg de erro
            buf = Right(buf, Len(buf) - lf) 'Cut the parsed data from the buffer
           'This is useful if theexternal hardware sends multiple barcodes in a time.
           
            cell_idx = cell_idx + 1 'increment the row number of cell where the barcode will be stored
            Cells(cell_idx, 1) = s
            Cells(cell_idx, 2) = Now()
            x = cell_idx

        Loop
        If Cells(5, 13) <> "" Then
            Sleep Cells(5, 13) * 1000
            mede
        End If
  End Select
End Sub



Sub connect()
 Dim ports As String
  If StrokeReader1.Connected = False Then
    ports = StrokeReader1.PortsAvailable
    MsgBox "Porta COM disponivel para conectar: " & ports
  
  End If
  If Cells(3, 13) <> "" Then
   
    x = 1
    StrokeReader1.port = Cells(3, 13) 'celula M3 qual a porta COM
    StrokeReader1.BaudRate = 4800
    StrokeReader1.Parity = NOPARITY
    StrokeReader1.StopBits = ONESTOPBIT
    StrokeReader1.DsrFlow = False
    StrokeReader1.CtsFlow = False
    StrokeReader1.DTR = False
    StrokeReader1.RTS = False
    StrokeReader1.Connected = True
    If StrokeReader1.Error Then
        Debug.Print StrokeReader1.ErrorDescription
    End If
  End If
End Sub

Sub zera()
  Dim x As String
  Dim y As String
  Dim t As String

  
  x = "Z" & Chr(13) & Chr(10) ' Z= 119 ou 151
  y = "Chr(13)"
  z = &HD
  t = vbCrLf

  
  StrokeReader1.send x
  StrokeReader1.send y
  StrokeReader1.send z
  
End Sub

Sub mede()

  Dim x As String
  Dim y As String
  Dim z As Byte
  If StrokeReader1.Connected = True Then
    x = "P" & Chr(13) & Chr(10)
    y = "Chr(13)"
    z = &HD
  
    StrokeReader1.send x
    StrokeReader1.send y
    StrokeReader1.send z
  
  End If
  If StrokeReader1.Connected = False Then
    MsgBox "Por favor, conecte-se a porta COM."
  
  End If
  
End Sub

Sub desconnect()
  StrokeReader1.Connected = False
  If StrokeReader1.Error Then
    Debug.Print StrokeReader1.ErrorDescription
  End If
End Sub


