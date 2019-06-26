Attribute VB_Name = "modGPIBAgilent"
Option Explicit


Dim AGI82357B As VisaComLib.FormattedIO488  'declara Instrumento GPIB


Sub Init_Instrument()


    With AGI82357B      'Programa el instrumento GPIB según se requiera.
        
        .WriteString "*RST"                     'Reset the counter
        
        ' set timeout large enough to sent all data
        .IO.Timeout = 10000
        
        .WriteString "*CLS"                     'Clear event registers and error queue
        .WriteString "*SRE 0"                   'Clear service request enable register
        .WriteString "*ESE 0"                   'Clear event status enable register
        .WriteString ":STAT:PRES"               'Preset enable registers and
                                                ' transition filters for operation and
                                                ' questionable status structures.
        .WriteString ":FORM:DATA ASC"           'Specify ASCII format for result query responses
        .WriteString ":ROSC:SOUR EXT"           'Use external oscillator
        .WriteString ":ROSC:EXT:CHEC OFF"       'turn off the automatic detection
        .WriteString ":DIAG:CAL:INT:AUTO OFF"   'Disable automatic interpolator calibration
        '.WriteString ":DISP:ENAB OFF"           'Disable display
        .WriteString ":HCOP:CONT OFF"           'Disable printing
        .WriteString ":CALC:MATH:STAT OFF"      'Disable post-processing (math)
        .WriteString ":CALC:LIM:STAT OFF"       'Disable post-processing (limit testing)
        .WriteString ":CALC:AVER:STAT OFF"      'Disable post-processing (statistics)
                   'Specify continuous measurements
        .WriteString "CONF:FREQ (@1)"           'Configure for frequency measurement
        .WriteString ":FUNC 'FREQ 1'"           'Select frequency
        .WriteString ":FREQ:ARM:STAR:SOUR IMM"
        .WriteString ":FREQ:ARM:STOP:SOUR TIM"
        .WriteString ":FREQ:ARM:STOP:TIM " '& Cells(4, 5).value
         .WriteString ":INIT:CONT ON"
              
    End With

End Sub



Sub MeasureFreq()
' Sends command to the 53131A to measure frequency

    Dim reply As String
    
       
    With AGI82357B
        
        .WriteString "READ:FREQ?"
        reply = .ReadString     ' return value in ASCII format
    End With


End Sub

Function SetIO(ioAddress As String) As Boolean
    ' set the I/O address to the text box in case the
    ' user changed it.
    ' bring up the input dialog and save any changes to the
    ' text box
    'Dim mgr As AgilentRMLib.SRMCls
    
    
    On Error GoTo ioError
    
    'ioAddress = Cells(1, 10)
    'ioAddress = InputBox("Ingrese la dirección I/O del AGI82357B", "Set IO address", ioAddress)

    If Len(ioAddress) > 5 Then
        'Set mgr = New AgilentRMLib.SRMCls
        'Set AGI82357B = New VisaComLib.FormattedIO488
        'Set AGI82357B.IO = mgr.Open(ioAddress)
        'Cells(1, 10) = ioAddress
    End If
    
    SetIO = True
    Exit Function
    
ioError:

    SetIO = False
    'MsgBox "Set IO error:" & vbCrLf & Err.Description

End Function

