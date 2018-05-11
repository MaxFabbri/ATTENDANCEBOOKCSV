VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPutPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IPutPay

Private mFromDate           As Date
Private mToDate             As Date

Private DebugOperation      As String
Private mNominativo         As String

' dati personalizzazione

Dim CurrDate                As Date

Dim fileHandler             As Integer

Dim FileName                As String

Dim outList()               As String

Dim personRow               As String
Dim calendarRow             As String

Private Type dataCode
    Code        As String
    quantity    As Double
    minute      As Boolean
End Type

Private Type dayData
    Day         As Integer
    dataCodes() As dataCode
    quantity    As Double
End Type

' per i totali mensili uso questo
Dim pay         As cPay
Dim pays        As cPays

' grezzo
Dim monthDataRaw()  As dayData

Dim nameInFile          As Boolean
Dim matriculaInFile     As Boolean

Dim codeList()          As String

Dim dailyTotalInFile    As Boolean
Dim monthlyTotalInFile  As Boolean

Dim separator           As String

Private Function IPutPay_AfterInitializeDip(CM As String) As Boolean

    IPutPay_AfterInitializeDip = True

End Function

Private Sub IPutPay_BuildOutPutFile(WebCaller As Boolean)


    On Error GoTo Pay_ERROR
    
    ' serve solamente per aprire il file
    
    DebugOperation = "Apertura file " & FileName
    'DebugOut DebugOperation
    
    fileHandler = FreeFile
    Open FileName For Output As fileHandler
    
    ' inserisce l'intestazione
    Print #fileHandler, FillSemicolon("Foglio presenze " & GetItalianMonthsName(Month(mFromDate)) & " " & Year(mFromDate))
    
Pay_EXIT:

    Exit Sub
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_BuildOutPutFile" & vbCrLf & "(" & DebugOperation & ")", , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select


End Sub

Private Function IPutPay_ChangeDay(oday As CDatoGiornaliero) As ePrctInfoStatus

    On Error GoTo Pay_ERROR
    
    Dim hour        As cOra
    
    'DebugOut "Function IPutPay_ChangeDay"
    
    For Each hour In oday.ORE.Contains(codeList).Group
        InsertCode hour.Causale, hour.ResolvePayQuantity(oday.Due), Not hour.IsDailyPayType And Not hour.Centesimi, Day(oday.Giorno)
        IPutPay_ChangeDay = Successful
    Next
    
Pay_EXIT:

    Exit Function
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_ChangeDay" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo, , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Function IPutPay_CheckPeriod() As Boolean

    IPutPay_CheckPeriod = oDip.isActive

End Function

Private Function IPutPay_DipSelection(ByVal dFromDate As Date, ByVal dToDate As Date, ByVal Name As String) As ADODB.Recordset

End Function

Private Property Let IPutPay_FromDate(ByVal RHS As Date)
    mFromDate = RHS
End Property

Private Property Get IPutPay_FromDate() As Date
    IPutPay_FromDate = mFromDate
End Property

Private Property Get IPutPay_FullCalendar() As Boolean
    IPutPay_FullCalendar = True
End Property

Private Function IPutPay_GetContext() As String
    If InDesign() Then
        'IPutPay_GetContext = "PEI"
    End If
End Function

Private Property Get IPutPay_GetFileList() As String()
    IPutPay_GetFileList = outList
End Property

Private Function IPutPay_GetPassword() As String
    If InDesign() Then
        'IPutPay_GetPassword = "ARTECHADMIN"
    End If
End Function

Private Function IPutPay_GetUser() As String
    If InDesign() Then
        'IPutPay_GetUser = "ARTECH"
    End If
End Function

Private Function IPutPay_Initialize(WebCaller As Boolean) As Boolean

    On Error GoTo Pay_ERROR
    
    'DebugOut "Sub IPutPay_Initialize"
    
    With oDip
    
        .Initialize
        
        .FromDate = mFromDate
        .ToDate = mToDate
        
        .ReadsHours = True
        .ReadsDailyCalendar = True
        .ReadsDailyProfiles = True
        
    End With
    
    Dim location As eINIWhere

    location = eINIWhere.LocalINI
    If WebCaller Then
        location = GlobalINI
    End If
        
    ' mettere un percorso univoco per ogni CLIENTE
    FileName = RINI("Files", "Path", location)
    
    ' fare il push dei vari file aperti
    Push outList(), FileName
    
    calendarRow = separator & Format$(Day(mFromDate), "00") & " " & Left$(GetItalianWeekDaysName(Weekday(mFromDate, vbMonday)), 2)
    For CurrDate = mFromDate + 1 To mToDate
        calendarRow = Concat(calendarRow, Format$(Day(CurrDate), "00") & " " & Left$(GetItalianWeekDaysName(Weekday(CurrDate, vbMonday)), 2), separator)
    Next
    calendarRow = FillSemicolon(calendarRow)
    
    IPutPay_Initialize = True
    
Pay_EXIT:
    Exit Function
    

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_Initialize" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf, , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Function IPutPay_InitializeDip(CM As String) As Boolean

Dim Index   As Integer

    On Error GoTo Pay_ERROR

    'DebugOut "Sub IPutPay_InitializeDip"
    
    ' primo loop su causali mese selezionato
    Set pays = Nothing
    Set pays = New cPays
    
    mNominativo = oDip.Nominativo
    personRow = ""
    
    If (nameInFile) Then
        personRow = mNominativo
    End If
    
    If (matriculaInFile) Then
        personRow = Concat(personRow, oDip.Matricola, " ")
    End If
    
    personRow = FillSemicolon(personRow)
    
    IPutPay_InitializeDip = True
    
Pay_EXIT:
    Exit Function
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_InitializeDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo, , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Function


Private Sub IPutPay_InitializePerc(oPerc As IPercentageInfo)

    'DebugOut "Sub IPutPay_InitializePerc"
    
    With oPerc
        .DefaultCalculatedValue = Information
        If RINI("Opzioni", "Numero secondi di attesa al termine") = "" Then
            .WaitSec = 60
        Else
            .WaitSec = Val(RINI("Opzioni", "Numero secondi di attesa al termine"))
        End If
        
    End With

End Sub


Private Sub IPutPay_MakeSectionsKeys()

    Dim oS  As CSection

    On Error GoTo Pay_ERROR
    
    oSects.IniServerName = App.EXEName & ".ini"
    oSects.DescrizioneGenerale = App.FileDescription & StringFormat(" - Versione {0}.{1}.{2}", App.Major, App.Minor, App.Revision)
    
    Set oS = oSects.Add("Anagrafica")
    oS.AddKey "Nominativo Su File", ComboBox, "Da abilitare se si vuole inserire il nominativo sul file", "0,1", "Sì,No"
    oS.AddKey "Codice Matricola Su File", ComboBox, "Da abilitare se si vuole inserire la matricola sul file", "0,1", "Sì,No"
    
    Set oS = oSects.Add("Causali")
    oS.AddKey "Lista Codici Su File", ListPresentCode, "Elencare i codici delle causali da inserire sul file"
    
    Set oS = oSects.Add("Parametri")
    oS.AddKey "Totali Giornalieri Su File", ComboBox, "Da abilitare se si vogliono inserire i totali giornalieri su file", "0,1", "Sì,No"
    oS.AddKey "Totali Mensili Su File", ComboBox, "Da abilitare se si vogliono inserire i totali mensili su file", "0,1", "Sì,No"
    oS.AddKey "Separatore", StringText, "Inserisci il separatore ; o ,"
    
    
Pay_EXIT:
    Exit Sub
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_MakeSectionsKeys", , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

        
End Sub

Private Property Get IPutPay_MemoLogin() As Boolean

End Property

Private Function IPutPay_PreInitialize() As Boolean
    IPutPay_PreInitialize = True
End Function

Private Function IPutPay_PreInitializeDip(CM As String) As Boolean
    IPutPay_PreInitializeDip = True
End Function

Private Sub IPutPay_ReadINIParameters()

    On Error GoTo Pay_ERROR
    
    ' è il contrario 0 abilitato 1 disabilitato
    nameInFile = (Val(RINI("Anagrafica", "Nominativo Su File")) = 0)
    matriculaInFile = (Val(RINI("Anagrafica", "Codice Matricola Su File")) = 0)
    
    codeList = Split(RINI("Causali", "Lista Codici Su File"), ",")
    
    dailyTotalInFile = (Val(RINI("Parametri", "Totali Giornalieri Su File")) = 0)
    monthlyTotalInFile = (Val(RINI("Parametri", "Totali Mensili Su File")) = 0)
    
    separator = RINI("Parametri", "Separatore")
    If (separator = "") Then
        separator = ";"
    End If
    
    
Pay_EXIT:

    Exit Sub
    

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_Initialize" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf, , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select


End Sub

Private Property Get IPutPay_SilentMode() As Boolean

End Property

Private Sub IPutPay_Terminate(WebCaller As Boolean)

Dim oFileNotice As cFilePath

    On Error GoTo Pay_ERROR
    
    'DebugOut "Sub IPutPay_Terminate"
    
    Close
    
    If Not WebCaller Then
    
        Set oFileNotice = New cFilePath
        
        With oFileNotice
        
            .AddFile FileName, True
            
            If RINI("Opzioni", "Numero secondi di attesa al termine") = "" Then
                .WaitSec = 60
            Else
                .WaitSec = Val(RINI("Opzioni", "Numero secondi di attesa al termine"))
            End If
    
            If .IsReady Then
                .Show
            End If
        End With
            
        Set oFileNotice = Nothing
        
    End If
    
Pay_EXIT:
    Exit Sub

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_Terminate" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo, , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Sub


Private Function IPutPay_TerminateDip(eState As ePrctInfoStatus, CM As String) As ePrctInfoStatus

    On Error GoTo Pay_ERROR
    
    'DebugOut "Function IPutPay_TerminateDip()"
    
    If (pays.Count = 0) Then
        IPutPay_TerminateDip = ePrctInfoStatus.Failed
        Exit Function
    End If
    
    ' scrive intestazione e calendario
    Print #fileHandler, personRow
    Print #fileHandler, calendarRow
    
    Dim codeIndex As Integer
    ' loop sulle causali ordinate
    For codeIndex = 0 To UBound(codeList)
        ' cerca se la causale è stata movimentata nel mese
        If (Not pays.isCodeLoaded(codeList(codeIndex)) Is Nothing) Then
            ' prepara la riga
            Print #fileHandler, FillSemicolon(makeCodeRow(codeList(codeIndex)))
        End If
    Next
    
    If dailyTotalInFile Then
        Print #fileHandler, FillSemicolon(makeTotalDailyRow)
    End If
    
    BlankData
    
    IPutPay_TerminateDip = eState

Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_TerminateDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo, , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Function

Private Property Let IPutPay_ToDate(ByVal RHS As Date)
    mToDate = RHS
End Property

Private Property Get IPutPay_ToDate() As Date
    IPutPay_ToDate = mToDate
End Property

Private Function IPutPay_ViewCalendar(dFromDate As Date, dToDate As Date, ByVal Name As String) As Boolean

End Function

Private Function IPutPay_ViewFullCalendar(StartDate As Date, EndDate As Date, ByVal Name As String) As ADODB.Recordset


    On Error GoTo Pay_ERROR
    
    'DebugOut "IPutPay_ViewFullCalendar()"
    DebugOperation = "Visualizzazione calendario"
    
    Dim oC As New cCalendar
    With oC
        .webMode = True
        .NumberFileOutPut = 1
        .FileName(1) = "Nome file paghe"
        .CalendarMode = PreviousMonth
        .Caption = Name
        .Fullmonth = True
        .DynamicPeriod = True
        .FilterOutPutFile = "File CVS (*.CSV;*.*)|*.CSV;*.*"
        Set .Icon = LoadResPicture(101, 1)
        Set IPutPay_ViewFullCalendar = .Run()
        StartDate = .StartDate
        EndDate = .EndDate
    End With
    Set oC = Nothing
    
    mFromDate = StartDate
    mToDate = EndDate
    
Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_ViewFullCalendar" & vbCrLf & "(" & DebugOperation & ")", , gwebMode)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    

End Function

Private Property Get IPutPay_ViewLOG(LogList As String, Optional WaitSec As Integer = 15) As Boolean

End Property

Private Property Get IPutPay_ViewPercentage() As Boolean
    IPutPay_ViewPercentage = True
End Property


Private Function FillSemicolon(ByVal Row As String)

    ' conta i ; e li filla se non sono 32

    Dim counter As Integer
    counter = CountCharInString(Row, separator)
    If (counter < 32) Then
        Row = Row & String$(32 - counter, separator)
    End If
    
    FillSemicolon = Row
    
End Function

Private Sub InsertCode(Causale As String, quantity As Double, Minutes As Boolean, Day As Integer)
    
    On Error Resume Next
    
    If (UBound(monthDataRaw) < Day) Then
        ReDim Preserve monthDataRaw(1 To DaysOfMonth(mFromDate))
    End If
    
    Dim tmp As String
    tmp = monthDataRaw(Day).dataCodes(1).Code
    If (Err.Number <> 0) Then
        ReDim Preserve monthDataRaw(Day).dataCodes(1 To 1)
    End If
    
    On Error GoTo CatchError
    
    Dim Index   As Integer
    
    Set pay = pays.isCodeLoaded(Causale)
    If (pay Is Nothing) Then
        Set pay = pays.Add(Causale)
    End If
    With pay
        .Value = .Value + quantity
        .IsDaily = Not Minutes
    End With
        
    For Index = LBound(monthDataRaw(Day).dataCodes) To UBound(monthDataRaw(Day).dataCodes)
        With monthDataRaw(Day).dataCodes(Index)
            If ((.Code = Causale) Or (.Code = "")) Then
                .minute = Minutes
                .quantity = quantity
                .Code = Causale
                monthDataRaw(Day).quantity = monthDataRaw(Day).quantity + quantity
                Exit Sub
            End If
            
        End With
    Next
    
    ReDim Preserve monthDataRaw(Day).dataCodes(1 To Index)
    With monthDataRaw(Day).dataCodes(Index)
        .minute = Minutes
        .quantity = quantity
        .Code = Causale
    End With
    monthDataRaw(Day).quantity = monthDataRaw(Day).quantity + quantity

EXIT_Sub:
    Exit Sub

CatchError:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".InsertCode" & vbCrLf & "(" & DebugOperation & ")", , gwebMode)
    Case vbAbort
        Resume EXIT_Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    

End Sub

Private Function makeTotalDailyRow()

    Dim Row     As String
    Dim ses     As Boolean
    Dim D       As Date
    Dim try     As String
    
    'Row = separator
    
    For D = mFromDate To mToDate
        On Error Resume Next
        With monthDataRaw(Day(D))
            try = .dataCodes(1).Code
            If (Err.Number = 0) Then
                On Error GoTo 0
                Dim Index   As Integer
                For Index = LBound(.dataCodes) To UBound(.dataCodes)
                    If .dataCodes(Index).minute Then
                        ' verifica se nella giornata c'è una voce sessagesimale
                        ses = True
                    End If
                Next
                If (ses) Then
                    If (Row = "") Then
                        Row = separator & MinuteToTime(.quantity, , "hh.nn")
                    Else
                        Row = Concat(Row, MinuteToTime(.quantity, , "hh.nn"), separator)
                    End If
                Else
                    If (Row = "") Then
                        If (.quantity > 0) Then
                            Row = separator & .quantity
                        Else
                            Row = separator
                        End If
                    Else
                        Row = Concat(Row, .quantity, separator)
                    End If
                End If
            End If
            If (Err.Number <> 0) Then
                If (Row = "") Then
                    Row = separator '& Row & separator
                Else
                    Row = Row & separator
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End With
    Next
    
    makeTotalDailyRow = Row

End Function

Private Function makeCodeRow(Code As String)

    Dim Row     As String
    Dim D       As Date
    Dim try     As String
    Dim Found   As Boolean
    
    Dim tot     As Double
    Dim ses     As Boolean
    
    Row = Code
    
    For D = mFromDate To mToDate
        Found = False
        On Error Resume Next
        With monthDataRaw(Day(D))
            try = .dataCodes(1).Code
            If (Err.Number = 0) Then
                On Error GoTo 0
                Dim Index   As Integer
                For Index = LBound(.dataCodes) To UBound(.dataCodes)
                    'ses = False
                    If (.dataCodes(Index).Code = Code) Then
                        Found = True
                        If .Day Then
                            ' non bisogna trasformarlo
                            Row = Concat(Row, .dataCodes(Index).quantity, separator)
                        Else
                            ' è in sessantesimi
                            Row = Concat(Row, MinuteToTime(.dataCodes(Index).quantity, False, "hh.nn"), separator)
                            ses = True
                        End If
                        ' trovato nella giornata
                        tot = tot + .dataCodes(Index).quantity
                    End If
                Next
            End If
            Err.Clear
            On Error GoTo 0
            If (Not Found) Then
                ' inserisce ,
                Row = Row & separator
            End If
        End With
    Next
    
    ' al termine inserisce il totale mensile
    If monthlyTotalInFile Then
        If (ses) Then
            Row = Concat(Row, MinuteToTime(tot, False, "hh.nn"), separator)
        Else
            Row = Concat(Row, tot, separator)
        End If
    End If
    
    makeCodeRow = Row

End Function


Private Sub BlankData()

    Dim D       As Date
    Dim try     As String
    
    For D = mFromDate To mToDate
        On Error Resume Next
        With monthDataRaw(Day(D))
            try = .dataCodes(1).Code
            If (Err.Number = 0) Then
                Dim Index   As Integer
                For Index = LBound(.dataCodes) To UBound(.dataCodes)
                    .dataCodes(Index).quantity = 0
                    .dataCodes(Index).Code = ""
                    .dataCodes(Index).minute = False
                Next
            End If
            .quantity = 0
        End With
        Err.Clear
        On Error GoTo 0
    Next

End Sub


