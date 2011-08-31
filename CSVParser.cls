VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private f As Integer
Private Buffer As String
Private QuoteMode As Boolean
Private CRMode As Boolean
Private TokenOpen As Boolean
Private Token As String
Private Tokens As Collection
Private EndOfFile As Boolean
Private FileIsOpen As Boolean
Private TokensBuffer As Collection
Private LineClosed As Boolean

Private Sub Class_Initialize()
    FileIsOpen = False
    EndOfFile = True
End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub

Function OpenFile(Filename)
    CloseFile
    f = FreeFile
    Open Filename For Input Access Read As f
    FileIsOpen = True
    EndOfFile = False
    QuoteMode = False
    CRMode = False
    TokenOpen = True
    LineClosed = False
    Set Tokens = New Collection
    Set TokensBuffer = Nothing
End Function

Function CloseFile()
    If FileIsOpen Then
        Close f
        FileIsOpen = False
    End If
    EndOfFile = True
    Set Tokens = Nothing
End Function

Property Get EndOfData() As Boolean
    EndOfData = EndOfFile
End Function

Function ReadFields() As Collection
    Set ReadFields = Nothing
    If Not FileIsOpen Then
        Exit Function
    End If

    Do Until EOF(f)
        Buffer = Buffer & Input(1, f)
        If QuoteMode Then
            Call QuoteModeInput
        Else
            Call NormalModeInput
        End If
        If LineClosed Then
            Set ReadFields = TokensBuffer
            LineClosed = False
            Exit Function
        End If
    Loop
    If TokenOpen Then
        Call CloseLine
        Set ReadFields = TokensBuffer
        Set TokensBuffer = Nothing
    End If
    CloseFile
End Function

Private Sub NormalModeInput()
    If CRMode And Buffer = vbLf Then
        CRMode = False
        Buffer = ""
        Exit Sub
    End If
    
    Select Case Buffer
    Case vbCr
        Call CloseLine
        CRMode = True
    Case vbLf
        Call CloseLine
    Case ","
        Call CloseToken
        TokenOpen = True
    Case """"
        If Token = "" Then
            QuoteMode = True
            Buffer = ""
        Else
            Call AddToken
        End If
    Case Else
        Call AddToken
    End Select
End Sub

Private Sub QuoteModeInput()
    Select Case Buffer
    Case """"
        'Do nothing
    Case """"""
        Buffer = """"
        Call AddToken
    Case """"""
        Buffer = """"
        Call AddToken
    Case ""","
        Buffer = ""
        Call CloseToken
        TokenOpen = True
    Case """" & vbCr
        Call CloseLine
        CRMode = True
    Case """" & vbLf
        Call CloseLine
    Case Else
        Call AddToken
    End Select
End Sub

Private Sub AddToken()
    Token = Token & Buffer
    Buffer = ""
End Sub

Private Sub CloseToken()
    QuoteMode = False
    Tokens.Add Token
    Token = ""
    Buffer = ""
    TokenOpen = False
End Sub

Private Sub CloseLine()
    Call CloseToken
    Set TokensBuffer = Tokens
    LineClosed = True
    TokenOpen = True
    Set Tokens = New Collection
End Sub
