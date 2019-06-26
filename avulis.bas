Attribute VB_Name = "modAvilus"
'== ADVANTEST ULIS Series Spectrum Analyzers Include File ==================
'===========================================================================
' Instrument Specific Error Codes
'===========================================================================
Global Const VI_ERROR_INSTR_FILE_OPEN = &HBFFC0800
Global Const VI_ERROR_INSTR_FILE_WRITE = &HBFFC0801

Global Const VI_ERROR_INSTR_INTERPRETING_RESPONSE = &HBFFC0803
Global Const VI_ERROR_INSTR_PARAMETER9 = &HBFFC0809
Global Const VI_ERROR_INSTR_PARAMETER10 = &HBFFC080A
Global Const VI_ERROR_INSTR_PARAMETER11 = &HBFFC080B
Global Const VI_ERROR_INSTR_PARAMETER12 = &HBFFC080C
Global Const VI_ERROR_INSTR_PARAMETER13 = &HBFFC080D
Global Const VI_ERROR_INSTR_PARAMETER14 = &HBFFC080E
Global Const VI_ERROR_INSTR_PARAMETER15 = &HBFFC080F

Global Const PREFIX_ERROR_INVALID_CONFIGURATION = &HBFFC09F0

Global Const VI_ERROR_MARKERINACTIVE = &HBFFC0810

'===========================================================================
' CONSTANTS FOR instrFeatures STRUCTURE
'===========================================================================
Global Const FREQUENCYLOW = 1&
Global Const FREQUENCYHIGH = 2&
Global Const SPANLOW = 3&
Global Const SPANHIGH = 4&
Global Const FREQOFFSETMIN = 5&
Global Const FREQOFFSETMAX = 6&
Global Const INPUTATTMIN = 7&
Global Const INPUTATTMAX = 8&
Global Const REFLEVELMIN = 9&
Global Const REFLEVELMAX = 10&
Global Const REFLEVELOFFSETMIN = 11&
Global Const REFLEVELOFFSETMAX = 12&
Global Const SWEEPTIMEMINI = 13&
Global Const SWEEPTIMEMAXI = 14&
Global Const GATEPOSITIONMINI = 15&
Global Const GATEPOSITIONMAXI = 16&
Global Const GATEWIDTHMINI = 17&
Global Const GATEWIDTHMAXI = 18&
Global Const AVERAGECOUNTMINI = 19&
Global Const AVERAGECOUNTMAXI = 20&
Global Const DLLEVELMIN = 21&
Global Const DLLEVELMAX = 22&
Global Const DISPLAYMODE = 23&
Global Const REFUNIT = 24&
Global Const SPAN = 25&
Global Const DELTAYMIN = 27&
Global Const DELTAYMAX = 28&
Global Const CALLEVELMIN = 29&
Global Const CALLEVELMAX = 30&
Global Const MKRMstatus = 31&


'===========================================================================
' GLOBAL USER-CALLABLE FUNCTION DECLARATIONS (Exportable Functions)
'===========================================================================
Declare Function AVULIS_init Lib "AVULIS_32.dll" (ByVal resourceName As String, ByVal IDquery As Integer, ByVal resetDevice As Integer, instrID As Long) As Long
Declare Function AVULIS_configRead Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal trace As Integer, ByVal peakSearch As Integer, traceAccuracy As Integer, traceArray As Integer, sweepTime As Double, ByVal sweepUnit As String, MKRNposition As Double, MKRNlevel As Double) As Long

'==============================
'   Configurations Fonctions
'==============================

' Configure Frequency
Declare Function AVULIS_Set_Freq Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal frequencyType As Integer, ByVal frequencyValue As Double) As Long
Declare Function AVULIS_Set_Span Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal spanType As Integer, ByVal spanValue As Double) As Long
Declare Function AVULIS_Set_FreqOffset Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal activation As Integer, ByVal frequencyOffset As Double) As Long
Declare Function AVULIS_Set_FreqStepSize Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal controlType As Integer, ByVal frequencyStep As Double) As Long
Declare Function AVULIS_Set_FreqCounter Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal activation As Integer, ByVal resolution As Integer) As Long

' Configure Amplitude
Declare Function AVULIS_Set_RefLevel Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal refLevelValue As Double) As Long
Declare Function AVULIS_Set_RefUnit Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal referenceUnit As Integer) As Long
Declare Function AVULIS_Set_Scale Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal scaleYaxis As Integer) As Long
Declare Function AVULIS_Set_RefOffset Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal activation As Integer, ByVal offset As Double) As Long
Declare Function AVULIS_Set_Att Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal controlType As Integer, ByVal value As Double) As Long
Declare Function AVULIS_Set_HighSenseMode Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal mode As Integer) As Long

' Configure Bandwidth
Declare Function AVULIS_Set_RBW Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal RBWSetting As Integer) As Long
Declare Function AVULIS_Set_VBW Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal VBWSetting As Integer) As Long
Declare Function AVULIS_Set_All_Auto_Couple Lib "AVULIS_32.dll" (ByVal instrID As Long) As Long


' Configure Sweep
Declare Function AVULIS_Set_SweepTime Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal controlType As Integer, ByVal sweepTime As Double) As Long
Declare Function AVULIS_Set_SweepMode Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal sweepMode As Integer, ByVal windowSweep As Integer) As Long
Declare Function AVULIS_Set_Trigger Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal source As Integer, ByVal polarity As Integer, ByVal videoLevel As Long, ByVal externalLevel As Double) As Long
Declare Function AVULIS_Set_TriggerPosition Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal position As Integer) As Long
Declare Function AVULIS_Set_TriggerTV Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal triggerSource As Integer, ByVal triggerPolarity As Integer, ByVal lineNumber As Integer) As Long

' Configure Trace
Declare Function AVULIS_Set_TraceAMode Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal traceAMode As Integer, ByVal averageCount As Integer, ByVal averageMode As Integer) As Long
Declare Function AVULIS_Set_TraceBMode Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal traceBMode As Integer) As Long
Declare Function AVULIS_Set_TraceDetector Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal detectorMode As Integer) As Long

' Configure Display & Limit Lines
Declare Function AVULIS_Set_LimitLine Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal limitLine As Integer, ByVal LL_FreqTable As String, ByVal LL_LevelTable As String) As Long
Declare Function AVULIS_Activate_LimitLine Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal LL1Status As Integer, ByVal LL2Status As Integer, ByVal levelXaxis As Integer, ByVal levelYaxis As Integer, ByVal passfail As Integer) As Long
Declare Function AVULIS_Set_DisplayLine Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal showDL As Integer, ByVal DLposition As Double) As Long

' Configure Measurement Window
Declare Function AVULIS_Set_MeasurementWindow Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal windowStatus As Integer, ByVal Xposition As Double, ByVal Xwidth As Double, ByVal Yposition As Double, ByVal Ywidth As Double) As Long

' Configure Markers Normal & Delta
Declare Function AVULIS_Set_NormalMarker Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal position As Double) As Long
Declare Function AVULIS_Set_DeltaMarker Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal position As Double) As Long
Declare Function AVULIS_Set_DeltaMarkerOptions Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal divided As Integer, ByVal fixed As Integer) As Long
Declare Function AVULIS_Set_MarkerTrace Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal markerTrace As Integer) As Long

' Configure Multimarkers
Declare Function AVULIS_Activate_Multimarkers Lib "AVULIS_32.dll" (ByVal instrID As Long) As Long
Declare Function AVULIS_Set_Multimarkers Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal MKRMstatus As String, ByVal MKRMpositions As String) As Long
Declare Function AVULIS_Set_ActiveMultiMarker Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal MKRMactive As Integer) As Long
Declare Function AVULIS_Delete_Markers Lib "AVULIS_32.dll" (ByVal instrID As Long) As Long

' Configure Peak Search
Declare Function AVULIS_Set_PeakSearchOptions Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal continuousSearch As Integer, ByVal signalTrack As Integer, ByVal peakRange As Integer, ByVal deltaY As Double) As Long

'=============================
'   Action/Status Functions
'=============================

Declare Function AVULIS_Search_Peak Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal searchType As Integer) As Long
Declare Function AVULIS_Execute_PeakTo Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal peakTo As Integer) As Long
Declare Function AVULIS_Set_MarkerDeltaTo Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal markerDeltaTo As Integer) As Long

' Trace Action
Declare Function AVULIS_Take_Sweep Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal startSweep As Integer) As Long
Declare Function AVULIS_Execute_TraceMath Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal traceMath As Integer) As Long
Declare Function AVULIS_NormalizingMode Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal mode As Integer) As Long
Declare Function AVULIS_Control_TraceAverage Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal traceAverage As Integer) As Long

' Configuration Read
Declare Function AVULIS_Read_FreqSettings Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal frequencyParameter As Integer, frequencyValue As Double) As Long
Declare Function AVULIS_Read_FreqOffset Lib "AVULIS_32.dll" (ByVal instrID As Long, status As Integer, value As Double) As Long
Declare Function AVULIS_Read_FreqStepSize Lib "AVULIS_32.dll" (ByVal instrID As Long, frequencyStepSize As Double, controlStatus As Integer) As Long
Declare Function AVULIS_Read_FreqCounter Lib "AVULIS_32.dll" (ByVal instrID As Long, counterValue As Double, counterStatus As Integer) As Long
Declare Function AVULIS_Read_ReferenceLevel Lib "AVULIS_32.dll" (ByVal instrID As Long, refLevelValue As Double) As Long
Declare Function AVULIS_Read_RefLevelOffset Lib "AVULIS_32.dll" (ByVal instrID As Long, value As Double, controlType As Integer) As Long
Declare Function AVULIS_Read_RefUnit Lib "AVULIS_32.dll" (ByVal instrID As Long, referenceUnitNumber As Integer, ByVal referenceUnitName As String) As Long
Declare Function AVULIS_Read_Scale Lib "AVULIS_32.dll" (ByVal instrID As Long, status As Integer, ByVal scaleName As String) As Long
Declare Function AVULIS_Read_Attenuation Lib "AVULIS_32.dll" (ByVal instrID As Long, valuedB As Double, controlType As Integer) As Long
Declare Function AVULIS_Read_RBWandVBW Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal bandwidthtoRead As Integer, bandwidthValue As Double, controlType As Integer) As Long
Declare Function AVULIS_Read_SweepTime Lib "AVULIS_32.dll" (ByVal instrID As Long, timeValue As Double, controlType As Integer) As Long
Declare Function AVULIS_Read_SweepMode Lib "AVULIS_32.dll" (ByVal instrID As Long, sweepModeStatus As Integer, ByVal sweepMode As String) As Long
Declare Function AVULIS_Read_Trigger Lib "AVULIS_32.dll" (ByVal instrID As Long, triggerStatus As Integer, ByVal triggerName As String, triggerFeature As Double) As Long
Declare Function AVULIS_Read_TriggerPosition Lib "AVULIS_32.dll" (ByVal instrID As Long, triggerPosition As Integer) As Long
Declare Function AVULIS_Read_DisplayLine Lib "AVULIS_32.dll" (ByVal instrID As Long, level As Double, status As Integer, ByVal message As String) As Long
Declare Function AVULIS_Read_LimitLineOptions Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal parameter As Integer, levelStatus As Integer, ByVal levelMessage As String) As Long
Declare Function AVULIS_Read_MeasurementWindow Lib "AVULIS_32.dll" (ByVal instrID As Long, status As Integer, ByVal message As String, Xposition As Double, Xwidth As Double, Yposition As Double, Ywidth As Double) As Long
Declare Function AVULIS_Read_Peak_Options Lib "AVULIS_32.dll" (ByVal instrID As Long, contPeakStatus As Integer, ByVal contPeakMessage As String, signalTrackStatus As Integer, ByVal signalTrackMessage As String, deltaY As Double) As Long

' Status Read
Declare Function AVULIS_Read_TraceStatus Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal traceModetoRead As Integer, traceModeStatus As Integer, ByVal generalStatusText As String) As Long
Declare Function AVULIS_Read_DetectionMode Lib "AVULIS_32.dll" (ByVal instrID As Long, detectorModeStatus As Integer, ByVal detectorMode As String) As Long
Declare Function AVULIS_Read_LLjudgment Lib "AVULIS_32.dll" (ByVal instrID As Long, result As Integer, ByVal message As String) As Long
Declare Function AVULIS_Read_LimitLineStatus Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal limitLinetoCheck As Integer, limitLineStatus As Integer, ByVal limitLineMessage As String) As Long
Declare Function AVULIS_Read_MarkerStatus Lib "AVULIS_32.dll" (ByVal instrID As Long, markerStatus As Integer, ByVal markerStatusMessage As String) As Long
Declare Function AVULIS_Read_DeltaMarkerOptions Lib "AVULIS_32.dll" (ByVal instrID As Long, fixedMarkerStatus As Integer, ByVal fixedMarkerMessage As String, dividedMerkerStatus As Integer, dividedMarkerValue As Double) As Long
Declare Function AVULIS_Read_MarkerTraceStatus Lib "AVULIS_32.dll" (ByVal instrID As Long, markerTraceStatus As Integer, ByVal markerTraceMessage As String) As Long
Declare Function AVULIS_Read_MultimarkersStatus Lib "AVULIS_32.dll" (ByVal instrID As Long, multimarkersStatus As Integer, ByVal multimarkersMessage As String) As Long

'====================
'   Data Functions
'====================

' Trace Function
Declare Function AVULIS_Read_TraceData Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal tracetoRead As Integer, ByVal dataFormat As Integer, traceDataArray As Integer) As Long
Declare Function AVULIS_Set_TraceAccuracy Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal accuracy As Integer) As Long
Declare Function AVULIS_Read_TraceAccuracy Lib "AVULIS_32.dll" (ByVal instrID As Long, accuracy As Integer) As Long

' Markers Functions
Declare Function AVULIS_Read_NormalMarker Lib "AVULIS_32.dll" (ByVal instrID As Long, position As Double, level As Double) As Long
Declare Function AVULIS_Read_DeltaMarker Lib "AVULIS_32.dll" (ByVal instrID As Long, position As Double, level As Double) As Long
Declare Function AVULIS_Read_Multimarkers Lib "AVULIS_32.dll" (ByVal instrID As Long, MKRMpositions As Double, MKRDposition As Double, MKRMlevels As Double, MKRDlevel As Double) As Long
Declare Function AVULIS_Read_MultiMarkerActive Lib "AVULIS_32.dll" (ByVal instrID As Long, multiMarkerActive As Integer) As Long

'=======================
'   Utility Functions
'=======================

Declare Function AVULIS_Calibration Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal elementtoCalibrate As Integer, ByVal calibrationLevel As Double, ByVal calibrationCorrection As Integer, ByVal frequencyCorrection As Integer) As Long
Declare Function AVULIS_Display_Label Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal showLabel As Integer, ByVal label As String) As Long
Declare Function AVULIS_Read_StatusByte Lib "AVULIS_32.dll" (ByVal instrID As Long, statusByteRead As Integer) As Long
Declare Function AVULIS_Clear_StatusByte Lib "AVULIS_32.dll" (ByVal instrID As Long) As Long

' Frequency or Time Conversion
Declare Function AVULIS_Convert_Frequency Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal inputFrequencyValue As Double, outputFrequencyValue As Double, ByVal frequencyUnit As String) As Long
Declare Function AVULIS_Convert_Time Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal inputTimeValue As Double, outputTimeValue As Double, ByVal timeUnit As String) As Long

' Memory Card Management
Declare Function AVULIS_Init_MemoryCard Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal drive As Integer) As Long
Declare Function AVULIS_Select_MemoryCard Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal drive As Integer) As Long
Declare Function AVULIS_Copy_MemoryCard Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal parameter As Integer) As Long
Declare Function AVULIS_Recall_File Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal drive As Integer, ByVal fileName As String) As Long
Declare Function AVULIS_Save_File Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal drive As Integer, ByVal format As Integer, ByVal fileName As String, ByVal setup As Integer, ByVal trace As Integer, ByVal limit As Integer, ByVal normalize As Integer, ByVal antenna As Integer, ByVal IDlist As Integer) As Long

' Antenna Correction
Declare Function AVULIS_Set_TR_AntennaCorrection Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal antennaChoice As Integer) As Long
Declare Function AVULIS_Set_User_AntennaCorrection Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal antennaCorrection As Integer, ByVal mdcorrection As Integer, ByVal FreqTable As String, ByVal LevelTable As String) As Long
Declare Function AVULIS_Read_AntennaCorrectionStatus Lib "AVULIS_32.dll" (ByVal instrID As Long, status As Integer, ByVal message As String) As Long

' Standard Utility Functions
Declare Function AVULIS_WriteInstrData Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal writeBuffer As String) As Long
Declare Function AVULIS_readInstrData Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal numberBytesToRead As Long, ByVal readBuffer As String, numBytesRead As Long) As Long
Declare Function AVULIS_reset Lib "AVULIS_32.dll" (ByVal instrSession As Long) As Long
Declare Function AVULIS_self_test Lib "AVULIS_32.dll" (ByVal instrID As Long, selfTestResult As Integer, ByVal selfTestMessage As String) As Long
Declare Function AVULIS_error_query Lib "AVULIS_32.dll" (ByVal instrID As Long, errorCode As Long, ByVal errorMessage As String) As Long
Declare Function AVULIS_error_message Lib "AVULIS_32.dll" (ByVal instrID As Long, statusCode As Long, ByVal message As String) As Long
Declare Function AVULIS_revision_query Lib "AVULIS_32.dll" (ByVal instrID As Long, ByVal instrumentDriverRevision As String, ByVal firmwareRevision As String) As Long
Declare Function AVULIS_close Lib "AVULIS_32.dll" (ByVal instrID As Long) As Long

'===========================================================================
' END INCLUDE FILE
'===========================================================================
