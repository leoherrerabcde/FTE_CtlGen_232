Attribute VB_Name = "GPIB1"
Global Const VI_ERROR = &H80000000



' /*****************************************************************************/
' /*  hp837xx.bas                                                              */
' /*  Copyright (C) 1999 Hewlett-Packard Company                               */
' /*                                                                           */
' /*  Driver for hp837xx Synthesized Signal Generator                          */
' /*  Driver Version: A.01.00                                                  */
' /*****************************************************************************/
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*                            global constants                               */
' /*---------------------------------------------------------------------------*/
'
'
Global Const hp837xx_INSTR_ERROR_NULL_PTR = (VI_ERROR + &H3FFC0D02)
Global Const hp837xx_INSTR_ERROR_RESET_FAILED = (VI_ERROR + &H3FFC0D03)
Global Const hp837xx_INSTR_ERROR_UNEXPECTED = (VI_ERROR + &H3FFC0D04)
Global Const hp837xx_INSTR_ERROR_INV_SESSION = (VI_ERROR + &H3FFC0D05)
Global Const hp837xx_INSTR_ERROR_LOOKUP = (VI_ERROR + &H3FFC0D06)
Global Const hp837xx_INSTR_ERROR_DETECTED = (VI_ERROR + &H3FFC0D07)
Global Const hp837xx_INSTR_NO_LAST_COMMA = (VI_ERROR + &H3FFC0D08)
Global Const hp837xx_INSTR_INV_ASCII_NUMBER = (VI_ERROR + &H3FFC0D09)

Global Const hp837xx_INSTR_ERROR_PARAMETER9 = (VI_ERROR + &H3FFC0D20)
Global Const hp837xx_INSTR_ERROR_PARAMETER10 = (VI_ERROR + &H3FFC0D21)
Global Const hp837xx_INSTR_ERROR_PARAMETER11 = (VI_ERROR + &H3FFC0D22)
Global Const hp837xx_INSTR_ERROR_PARAMETER12 = (VI_ERROR + &H3FFC0D23)
Global Const hp837xx_INSTR_ERROR_PARAMETER13 = (VI_ERROR + &H3FFC0D24)
Global Const hp837xx_INSTR_ERROR_PARAMETER14 = (VI_ERROR + &H3FFC0D25)
Global Const hp837xx_INSTR_ERROR_PARAMETER15 = (VI_ERROR + &H3FFC0D26)
Global Const hp837xx_INSTR_ERROR_PARAMETER16 = (VI_ERROR + &H3FFC0D27)
Global Const hp837xx_INSTR_ERROR_PARAMETER17 = (VI_ERROR + &H3FFC0D28)
Global Const hp837xx_INSTR_ERROR_PARAMETER18 = (VI_ERROR + &H3FFC0D29)

'         /*----------------------------------------------------------------*
'          |  Constants used by system status functions.  These defines     |
'          |    are bit numbers which define the operation and questionable |
'          |    registers.                                                  |
'          *----------------------------------------------------------------*/
'
Global Const hp837xx_QUES_BIT0 = 1
Global Const hp837xx_QUES_BIT1 = 2
Global Const hp837xx_QUES_BIT2 = 4
Global Const hp837xx_QUES_BIT3 = 8
Global Const hp837xx_QUES_BIT4 = 16
Global Const hp837xx_QUES_BIT5 = 32
Global Const hp837xx_QUES_BIT6 = 64
Global Const hp837xx_QUES_BIT7 = 128
Global Const hp837xx_QUES_BIT8 = 256
Global Const hp837xx_QUES_BIT9 = 512
Global Const hp837xx_QUES_BIT10 = 1024
Global Const hp837xx_QUES_BIT11 = 2048
Global Const hp837xx_QUES_BIT12 = 4096
Global Const hp837xx_QUES_BIT13 = 8192
Global Const hp837xx_QUES_BIT14 = 16384
Global Const hp837xx_QUES_BIT15 = 32768
Global Const hp837xx_OPER_BIT0 = 1
Global Const hp837xx_OPER_BIT1 = 2
Global Const hp837xx_OPER_BIT2 = 4
Global Const hp837xx_OPER_BIT3 = 8
Global Const hp837xx_OPER_BIT4 = 16
Global Const hp837xx_OPER_BIT5 = 32
Global Const hp837xx_OPER_BIT6 = 64
Global Const hp837xx_OPER_BIT7 = 128
Global Const hp837xx_OPER_BIT8 = 256
Global Const hp837xx_OPER_BIT9 = 512
Global Const hp837xx_OPER_BIT10 = 1024
Global Const hp837xx_OPER_BIT11 = 2048
Global Const hp837xx_OPER_BIT12 = 4096
Global Const hp837xx_OPER_BIT13 = 8192
Global Const hp837xx_OPER_BIT14 = 16384
Global Const hp837xx_OPER_BIT15 = 32768
'         /*----------------------------------------------------------------*
'          |          Constants used by function hp837xx_timeOut             |
'          *----------------------------------------------------------------*/
'
Global Const hp837xx_TIMEOUT_MIN = 0&
Global Const hp837xx_TIMEOUT_MAX = 2147483647
'         /*----------------------------------------------------------------*
'          |                    Miscellaneous #define's                     |
'          *----------------------------------------------------------------*/
'
Global Const hp837xx_CMDSTRINGARR_Q_MIN = 2&
Global Const hp837xx_CMDSTRINGARR_Q_MAX = 2147483647
Global Const hp837xx_CMDINT16ARR_Q_MIN = 1&
Global Const hp837xx_CMDINT16ARR_Q_MAX = 2147483647
Global Const hp837xx_CMDINT32ARR_Q_MIN = 1&
Global Const hp837xx_CMDINT32ARR_Q_MAX = 2147483647
Global Const hp837xx_CMDREAL64ARR_Q_MIN = 1&
Global Const hp837xx_CMDREAL64ARR_Q_MAX = 2147483647
'         /*----------------------------------------------------------------*
'          |              Instrument specific #defines's                    |
'          *----------------------------------------------------------------*/
'
' //  Level correct tables
Global Const hp837xx_LEVEL_CORRECT_TABLE_1 = 0
Global Const hp837xx_LEVEL_CORRECT_TABLE_2 = 1
Global Const hp837xx_LEVEL_CORRECT_TABLE_3 = 2
Global Const hp837xx_LEVEL_CORRECT_TABLE_4 = 3
Global Const hp837xx_LEVEL_CORRECT_TABLE_LASTENUM = 4

' //  Power Leveling types
Global Const hp837xx_PWR_LEVELING_TYPE_INTERNAL = 0
Global Const hp837xx_PWR_LEVELING_TYPE_DIODE = 1
Global Const hp837xx_PWR_LEVELING_TYPE_PMETER = 2
Global Const hp837xx_PWR_LEVELING_TYPE_LASTENUM = 3

' //  AM type
Global Const hp837xx_AM_TYPE_EXP = 0
Global Const hp837xx_AM_TYPE_LIN = 1
Global Const hp837xx_AM_TYPE_LASTENUM = 2

' // AM Sensitivity
Global Const hp837xx_AM_SENS_PCT_30 = 0
Global Const hp837xx_AM_SENS_PCT_100 = 1
Global Const hp837xx_AM_SENS_PCT_LASTENUM = 2

' //  Modulation Waveforms
Global Const hp837xx_MOD_WAVEFORM_SINUSOID = 0
Global Const hp837xx_MOD_WAVEFORM_SQUARE = 1
Global Const hp837xx_MOD_WAVEFORM_TRIANGLE = 2
Global Const hp837xx_MOD_WAVEFORM_RAMP = 3
Global Const hp837xx_MOD_WAVEFORM_UNIFORM = 4
Global Const hp837xx_MOD_WAVEFORM_GAUSSIAN = 5
Global Const hp837xx_MOD_WAVEFORM_LASTENUM = 6

' //  Modulation Couplings
Global Const hp837xx_MOD_COUPLING_AC = 0
Global Const hp837xx_MOD_COUPLING_DC = 1
Global Const hp837xx_MOD_COUPLING_LASTENUM = 2

' //  Internal PM Range
Global Const hp837xx_PM_INT_RANGE_AUTO = 0
Global Const hp837xx_PM_INT_RANGE_LOW = 1
Global Const hp837xx_PM_INT_RANGE_HIGH = 2
Global Const hp837xx_PM_INT_RANGE_LASTENUM = 3

' // Pulse triggering source
Global Const hp837xx_PULSE_TRIG_INTERNAL = 0
Global Const hp837xx_PULSE_TRIG_EXTERNAL = 1
Global Const hp837xx_PULSE_TRIG_DOUBLET = 2
Global Const hp837xx_PULSE_TRIG_GATED = 3
Global Const hp837xx_PULSE_TRIG_LASTENUM = 4

' // External Pulse Polarity
Global Const hp837xx_PULSE_POL_NORMAL = 0
Global Const hp837xx_PULSE_POL_INVERTED = 1
Global Const hp837xx_PULSE_POL_LASTENUM = 2

' //  Frequency step directions
Global Const hp837xx_FREQUENCY_STEP_UP = 0
Global Const hp837xx_FREQUENCY_STEP_DOWN = 1
Global Const hp837xx_FREQUENCY_STEP_MAX = 2

' // Power step directions
Global Const hp837xx_POWER_STEP_UP = 0
Global Const hp837xx_POWER_STEP_DOWN = 1
Global Const hp837xx_POWER_STEP_MAX = 2

' //Attenuator hold
Global Const hp837xx_ATTENUATOR_HOLD_ON = 0
Global Const hp837xx_ATTENUATOR_HOLD_OFF = 1
Global Const hp837xx_ATTENUATOR_HOLD_ONCE = 2
Global Const hp837xx_ATTENUATOR_HOLD_MAX = 3

' /*---------------------------------------------------------------------------*/
' /*                          end constants                                    */
' /*                                                                           */
' /*****************************************************************************/
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*                        Function Prototypes                                */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
'
'         /*----------------------------------------------------------------*
'          |  VXI plug&play required functions                              |
'          *----------------------------------------------------------------*/
'
Declare Function hp837xx_init Lib "hp837xx_32.dll" _
        (ByVal resourceName As String, _
         ByVal IDQuery As Integer, _
         ByVal resetDevice As Integer, _
         ByRef instrumentHandle As Long) As Long
'
Declare Function hp837xx_close Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long) As Long
'
Declare Function hp837xx_reset Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long) As Long
'
Declare Function hp837xx_self_test Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef selfTestResult As Integer, _
         ByVal selfTestMessage As String) As Long
'
Declare Function hp837xx_error_query Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef errorCode As Long, _
         ByVal errorMessage As String) As Long
'
Declare Function hp837xx_error_message Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal statusCode As Long, _
         ByVal message As String) As Long
'
Declare Function hp837xx_revision_query Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal instrumentDriverRevision As String, _
         ByVal firmwareRevision As String) As Long
'
'         /*----------------------------------------------------------------*
'          |  HP standard utility functions                                 |
'          *----------------------------------------------------------------*/
'
Declare Function hp837xx_timeOut Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal setTimeOut As Long) As Long
'
Declare Function hp837xx_timeOut_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef timeout As Long) As Long
'
Declare Function hp837xx_errorQueryDetect Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal setErrorQueryDetect As Integer) As Long
'
Declare Function hp837xx_errorQueryDetect_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef errorQueryDetect As Integer) As Long
'
Declare Function hp837xx_dcl Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long) As Long
'
'         /*----------------------------------------------------------------*
'          |  HP standard status functions                                  |
'          *----------------------------------------------------------------*/
'
Declare Function hp837xx_opc_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef instrumentReady As Integer) As Long
'
Declare Function hp837xx_readStatusByte_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef statusByte As Integer) As Long
'
Declare Function hp837xx_operEvent_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef operationEventRegister As Long) As Long
'
Declare Function hp837xx_operCond_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef operationConditionRegister As Long) As Long
'
Declare Function hp837xx_quesEvent_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef questionableEventRegister As Long) As Long
'
Declare Function hp837xx_quesCond_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef questionableConditionRegister As Long) As Long
'
'         /*----------------------------------------------------------------*
'          |  HP standard command passthrough functions                     |
'          *----------------------------------------------------------------*/
'
Declare Function hp837xx_cmd Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal sendStringCommand As String) As Long
'
Declare Function hp837xx_cmdString_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal queryStringCommand As String, _
         ByRef stringSize As Long, _
         ByRef stringResult As String) As Long
'
Declare Function hp837xx_cmdInt Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal sendIntegerCommand As String, _
         ByVal sendInteger As Long) As Long
'
Declare Function hp837xx_cmdInt16_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal queryI16Command As String, _
         ByRef i16Result As Integer) As Long
'
Declare Function hp837xx_cmdInt32_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal queryI32Command As String, _
         ByRef i32Result As Long) As Long
'
Declare Function hp837xx_cmdInt16Arr_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal queryI16ArrayCommand As String, _
         ByVal i16ArraySize As Long, _
         ByRef i16ArrayResult As Integer, _
         ByRef i16ArrayCount As Long) As Long
'
Declare Function hp837xx_cmdInt32Arr_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal queryI32ArrayCommand As String, _
         ByVal i32ArraySize As Long, _
         ByRef i32ArrayResult As Long, _
         ByRef i32ArrayCount As Long) As Long
'
Declare Function hp837xx_cmdReal Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal sendRealCommand As String, _
         ByVal sendReal As Double) As Long
'
Declare Function hp837xx_cmdReal64_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal queryRealCommand As String, _
         ByRef realResult As Double) As Long
'
Declare Function hp837xx_cmdReal64Arr_Q Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal realArrayCommand As String, _
         ByVal realArraySize As Long, _
         ByRef realArrayResult As Double, _
         ByRef realArrayCount As Long) As Long
'
'
'         /*----------------------------------------------------------------*
'          |  INSTRUMENT SPECIFIC FUNCTIONS                                 |
'          |    Function prototypes for instrument specific functions.      |
'          *----------------------------------------------------------------*
'          |  DEVELOPER: Add function prototypes here.  Remember that       |
'          |             function prototypes must be consistent with        |
'          |             the driver's function panel prototypes.            |
'          *----------------------------------------------------------------*/
'
'         /*----------------------------------------------------------------*
'          |  Sub-system Functions                                          |
'          *----------------------------------------------------------------*/
' /************************************/
' /* Power Level Correction Functions */
' /************************************/
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_levelSetCorrectTable                                             */
' /*    This method selects a level correct table and then loads the frequency */
' /*    points and correction factors into that table.                         */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 levelCorrectSetTableType                                       */
' /*      Selects one of four level correct tables where level correct data    */
' /*      will be loaded.                                                      */
' /*                                                                           */
' /*      hp837xx_LEVEL_CORRECT_TABLE_1 - Level correct table one              */
' /*      hp837xx_LEVEL_CORRECT_TABLE_2 - Level correct table two              */
' /*      hp837xx_LEVEL_CORRECT_TABLE_3 - Level correct table three            */
' /*      hp837xx_LEVEL_CORRECT_TABLE_4 - Level correct table four             */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 levelCorrectTableSize                                          */
' /*      The number of values in the level correct table.  This value must be */
' /*      between 4 and 802.                                                   */
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViAReal64 levelCorrectArray                                            */
' /*      Sets the user frequency and level correction values.  These values   */
' /*      must be sent in frequency and level correction pairs.The input       */
' /*      frequency range is 1 GHz to 20 GHz for the HP 83711A/11B/31A/31B     */
' /*      and 0.01 GHz to 20 GHz for the 83712A/12B/32A/32B.  The level        */
' /*      correction range is -40 dB to 40 dB.                                 */
' /*                                                                           */
' /*      The synthesizer will sort the entered list by frequency              */
' /*      automatically.  An instrument preset has no effect on the user level */
' /*      correction data.                                                     */
' /*                                                                           */
' /*      The number of values in this array must be equal to the              */
' /*      levelCorrectTableSize.                                               */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_levelSetCorrectTable Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal levelCorrectSetTableType As Integer, _
     ByVal levelCorrectTableSize As Integer, _
     ByRef levelCorrectArray As Double) As Long
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_levelGetCorrectTable                                             */
' /*    This method retrieves the frequency points and correction factors from */
' /*    the selected level correct table.                                      */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 levelCorrectGetTableType                                       */
' /*      Selects one of four level correct tables from which level correct    */
' /*      data will be retrieved.                                              */
' /*                                                                           */
' /*      hp837xx_LEVEL_CORRECT_TABLE_1 - Level correct table one              */
' /*      hp837xx_LEVEL_CORRECT_TABLE_2 - Level correct table two              */
' /*      hp837xx_LEVEL_CORRECT_TABLE_3 - Level correct table three            */
' /*      hp837xx_LEVEL_CORRECT_TABLE_4 - Level correct table four             */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 levelCorrectMaxArraySize                                       */
' /*      This values specifies the maximum size of the                        */
' /*      "levelCorrectFreqArray_Q" and "levelCorrectFactorArray_Q" arrays.    */
' /*      The value should be large enough for the points in the table.  The   */
' /*      range is from 8 to 802.                                              */
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViAReal64 levelCorrectArray_Q                                          */
' /*      Returns the frequency points and correction factors pairs that       */
' /*      are currently loaded in the selected table.                          */
' /*                                                                           */
' /*  PARAMETER FIVE                                                           */
' /*    ViPInt16 levelCorrectArraySize_Q                                       */
' /*      This is the actual number of elements returned in the                */
' /*      "levelCorrectFreqArray_Q" and "levelCorrectFactorArray_Q" arrays.    */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_levelGetCorrectTable Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal levelCorrectGetTableType As Integer, _
     ByVal levelCorrectMaxArraySize As Integer, _
     ByRef levelCorrectArray_Q As Double, _
     ByRef levelCorrectArraySize_Q As Integer) As Long
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_levelSetCorrectState                                             */
' /*    This method sets the state of the selected level correct table.        */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 levelCorrectStateTableType                                     */
' /*      Selects one of four level correct tables that is used to correct     */
' /*      power at the synthesized RF OUTPUT connector.                        */
' /*                                                                           */
' /*      hp837xx_LEVEL_CORRECT_TABLE_1 - Level correct table one              */
' /*      hp837xx_LEVEL_CORRECT_TABLE_2 - Level correct table two              */
' /*      hp837xx_LEVEL_CORRECT_TABLE_3 - Level correct table three            */
' /*      hp837xx_LEVEL_CORRECT_TABLE_4 - Level correct table four             */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViBoolean levelCorrectState                                            */
' /*      Turns the level correction state on or off.  The default state is    */
' /*      off.                                                                 */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_levelSetCorrectState Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal levelCorrectStateTableType As Integer, _
     ByVal levelCorrectState As Integer) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_pwrSetLevelingType                                               */
' /*    This method selects the type of leveling for output power automatic    */
' /*    level control.                                                         */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 pwrLevelingType                                                */
' /*      The leveling type to be used.                                        */
' /*                                                                           */
' /*      hp837xx_PWR_LEVELING_TYPE_INTERNAL - Internal leveling.              */
' /*      hp837xx_PWR_LEVELING_TYPE_DIODE - External diode detector leveling.  */
' /*      hp837xx_PWR_LEVELING_TYPE_PMETER - External power meter leveling.    */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_pwrSetLevelingType Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal pwrLevelingType As Integer) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_pwrSetMeterInitLevel                                             */
' /*    This method sets the initial reading of the external power meter to    */
' /*    the synthesizer for user during external power meter leveling.         */
' /*                                                                           */
' /*    The power meter reading set with his method allows the synthesizer to  */
' /*    calculate the value of voltage present at the power meter recorder     */
' /*    output connector.                                                      */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViReal64 pwrMeterInitLevel                                             */
' /*      The initial reading of the external power meter to the synthesizer.  */
' /*      The allowable range is -120 dBm (-100 dBm for HP 83711A/12A/31A/32A) */
' /*      to +30 dBm when option 1E1 is installed or -15 dBm to +30 dBm if     */
' /*      option 1E1 is not installed.                                         */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_pwrSetMeterInitLevel Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal pwrMeterInitLevel As Double) As Long
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_pwrSetMeterAddress                                               */
' /*    This method changes the HP-IB address that the synthesizer uses when   */
' /*    communicating with an external power meter during the level correct    */
' /*    routine.  It does not set the address of the power meter.              */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 pwrMeterAddress                                                */
' /*      The HP-IB address of the external power meter.  The valid address    */
' /*      range is 00 to 30.  The default value is 13.                         */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_pwrSetMeterAddress Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal pwrMeterAddress As Integer) As Long
'
'
' /********************************/
' /* Modulation Control Functions */
' /********************************/
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modAMConfigInt                                                   */
' /*    This method configures and sets the state of the internal amplitude    */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    This method requires option 1E2 be installed and is only available for */
' /*    83731B and 83732B.                                                     */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modAMIntState                                                */
' /*      Turns the internal amplitude modulation on or off.                   */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 modAMIntType                                                   */
' /*      Sets the internal amplitude modulation type.                         */
' /*                                                                           */
' /*      hp837xx_AMPLITUDE_MOD_TYPE_EXP - Selects exponential amplitude       */
' /*                                       modulation.                         */
' /*      hp837xx_AMPLITUDE_MOD_TYPE_LIN - Selects linear amplitude modulation.*/
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViReal64 modAMIntDepth                                                 */
' /*      Selects the AM depth when in internal logarithmic or linear AM mode. */
' /*     In linear mode, the allowed range is 0 to 100% with 0.1% resolution.  */
' /*     In logarithmic mode, the allowed range is 0 to 60 dB with 0.01 dB     */
' /*     resolution.                                                           */
' /*                                                                           */
' /*      When the internal AM depth is set between 30 dB and 60 dB, the entry */
' /*      resolution is 0.01 dB; however, the hardware resolution might be     */
' /*      slightly greater than 0.01 dB.  The hardware resolution will always  */
' /*      be less than 0.015 dB.                                               */
' /*                                                                           */
' /*  PARAMETER FIVE                                                           */
' /*    ViReal64 modAMIntGenRate                                               */
' /*      Sets the internal AM modulation rate.  The allowable range for the   */
' /*      parameter is 0.5 Hz to 100 kHz with a resolution of 0.5 Hz.  The     */
' /*      default value is 5kHz.                                               */
' /*                                                                           */
' /*  PARAMETER SIX                                                            */
' /*    ViInt16 modAMIntWaveform                                               */
' /*      Selects the waveform of the internal AM modulation generator.        */
' /*                                                                           */
' /*      hp837xx_MOD_WAVEFORM_SINUSOID - Selects the sinusoidal waveform.     */
' /*      hp837xx_MOD_WAVEFORM_SQUARE - Selects the square waveform.           */
' /*      hp837xx_MOD_WAVEFORM_TRIANGLE - Selects the triangle waveform.       */
' /*      hp837xx_MOD_WAVEFORM_RAMP - Selects the ramp waveform.               */
' /*      hp837xx_MOD_WAVEFORM_UNIFORM - Selects the uniform waveform.         */
' /*      hp837xx_MOD_WAVEFORM_GAUSSIAN - Selects the gaussian noise waveform. */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modAMConfigInt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modAMIntState As Integer, _
     ByVal modAMIntType As Integer, _
     ByVal modAMIntDepth As Double, _
     ByVal modAMIntGenRate As Double, _
     ByVal modAMIntWaveform As Integer) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modAMConfigExt                                                   */
' /*    This method configures and sets the state of the external amplitude    */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    This method is only available for 83731B and 83732B.                   */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modAMExtState                                                */
' /*      Turns the external amplitude modulation on or off.  The default      */
' /*      state is off.                                                        */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 modAMExtType                                                   */
' /*      Sets the external amplitude modulation type.                         */
' /*                                                                           */
' /*      hp837xx_AMPLITUDE_MOD_TYPE_EXP - Selects exponential amplitude       */
' /*                                       modulation.                         */
' /*      hp837xx_AMPLITUDE_MOD_TYPE_LIN - Selects linear amplitude modulation.*/
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViReal64 modAMExtSensitivity                                           */
' /*      Sets the linear AM sensitivity.  The allowable range for the         */
' /*      parameter is 30%/Volt to 100%/Volt.  The default is 30%/Volt.        */
' /*                                                                           */
' /*      hp837xx_AMPLITUDE_MOD_SENS_PCT_30 - Sets to 30% / Volt               */
' /*      hp837xx_AMPLITUDE_MOD_SENS_PCT_100 - Sets to 100% / Volt             */
' /*                                                                           */
' /*      If the external modulation is set to exponential, the AM sensitivity */
' /*      will be set to -10 dB/Volt and this setting will not take effect.    */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modAMConfigExt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modAMExtState As Integer, _
     ByVal modAMExtType As Integer, _
     ByVal modAMExtSensitivity As Integer) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modFMConfigInt                                                   */
' /*    This method configures and sets the state of the internal frequency    */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    This method requires option 1E2 be installed and is only available for */
' /*    83731B and 83732B.                                                     */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modFMIntState                                                */
' /*      This method turns frequency modulation on or off.  When the          */
' /*      synthesizer is set to the preset state, frequency modulation is      */
' /*      turned off.                                                          */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViReal64 modFMIntDeviation                                             */
' /*      Sets the internal FM deviation.  The allowable range is 0 Hz to      */
' /*      10 MHz.                                                              */
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViReal64 modFMIntFreqRate                                              */
' /*      Sets the FM internal modulation rate.  The allowable range for the   */
' /*      parameter is 1kHz to 1 MHz with a resolution of 0.5 Hz.              */
' /*                                                                           */
' /*  PARAMETER FIVE                                                           */
' /*    ViInt16 modFMIntWaveform                                               */
' /*      Selects the waveform of the internal FM modulation generator.        */
' /*                                                                           */
' /*      hp837xx_MOD_WAVEFORM_SINUSOID - Selects the sinusoidal waveform.     */
' /*      hp837xx_MOD_WAVEFORM_SQUARE - Selects the square waveform.           */
' /*      hp837xx_MOD_WAVEFORM_TRIANGLE - Selects the triangle waveform.       */
' /*      hp837xx_MOD_WAVEFORM_RAMP - Selects the ramp waveform.               */
' /*      hp837xx_MOD_WAVEFORM_UNIFORM - Selects the uniform waveform.         */
' /*      hp837xx_MOD_WAVEFORM_GAUSSIAN - Selects the gaussian noise waveform. */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modFMConfigInt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modFMIntState As Integer, _
     ByVal modFMIntDeviation As Double, _
     ByVal modFMIntFreqRate As Double, _
     ByVal modFMIntWaveform As Integer) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modFMConfigExt                                                   */
' /*    This method configures and sets the state of the external frequency    */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    When the coupling is set to DC and the state set to off, the           */
' /*    synthesizer circuitry is configured so that the FM/PM IN connector     */
' /*    will accept a modulating signal with a minimum rate of 1 kHz.  When    */
' /*    the coupling is set to DC and the state to on, the FM/PM IN connector  */
' /*    will accept a modulating signal with a minimum rate of 0 Hz(DC).       */
' /*                                                                           */
' /*    This method is only available for 83731B and 83732B.                   */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modFMExtState                                                */
' /*      This method turns frequency modulation on or off.  When the          */
' /*      synthesizer is set to the preset state, frequency modulation is      */
' /*      turned off.                                                          */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 modFMExtCoupling                                               */
' /*      Selects either AC or DC coupling for the FM/PM IN connector.  The    */
' /*      default state of FM coupling is AC.                                  */
' /*                                                                           */
' /*      hp837xx_MOD_COUPLING_AC - Sets coupling to AC.                       */
' /*      hp837xx_MOD_COUPLING_DC - Sets coupling to DC                        */
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViReal64 modFMExtSensitivity                                           */
' /*      Sets the FM sensitivity.  FM sensitivity is coupled to the CW        */
' /*      frequency.  As a result, any entered value will automatically adjust */
' /*      to the closest preset value for any given CW frequency.  The         */
' /*      allowable range is shown in the table below.                         */
' /*                                                                           */
' /*      FREQUENCY           SENSITIVITY                                UNIT  */
' /*      1 GHz to 20 GHz     10, 5, 3, 1, 0.3, 0.1, 0.003               MHz/V */
' /*      256 MHz to < 1GHz   2500, 1250, 750, 250, 75, 25, 7.5          kHz/V */
' /*      64 MHz to < 256 MHz 625, 312, 187, 62.5, 18.7, 6.25, 1.87      kHz/V */
' /*      16 MHz to < 64 MHz  156, 78.1, 46.8, 15.6, 4.68, 1.56, 0.468   kHz/V */
' /*      10 MHz to < 16 MHz  78.1, 39.0, 23,4, 7.81, 2.34, 0.781, 0.234 kHz/V */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modFMConfigExt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modFMExtState As Integer, _
     ByVal modFMExtCoupling As Integer, _
     ByVal modFMExtSensitivity As Double) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modPMConfigInt                                                   */
' /*    This method configures and sets the state of the internal phase        */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    This method is only available on HP 83731B/32B models with Option 800  */
' /*    installed.                                                             */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modPMIntState                                                */
' /*      This method turns internal phase modulation on or off.  When the     */
' /*      synthesizer is set to the preset state, pulse modulation is turned   */
' /*      off.                                                                 */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViReal64 modPMIntDeviation                                             */
' /*      Sets the internal PM deviation.  The allowable range is 0 rads to    */
' /*      200 rads.                                                            */
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViReal64 modPMIntFreqRate                                              */
' /*      Sets the PM internal modulation rate.  The allowable range is .05 Hz */
' /*      to 1 MHz with a resolution of 0.5 Hz.                                */
' /*                                                                           */
' /*  PARAMETER FIVE                                                           */
' /*    ViInt16 modPMIntWaveform                                               */
' /*      Selects the waveform of the internal PM modulation generator.        */
' /*                                                                           */
' /*      hp837xx_MOD_WAVEFORM_SINUSOID - Selects the sinusoidal waveform.     */
' /*      hp837xx_MOD_WAVEFORM_SQUARE - Selects the square waveform.           */
' /*      hp837xx_MOD_WAVEFORM_TRIANGLE - Selects the triangle waveform.       */
' /*      hp837xx_MOD_WAVEFORM_RAMP - Selects the ramp waveform.               */
' /*      hp837xx_MOD_WAVEFORM_UNIFORM - Selects the uniform waveform.         */
' /*      hp837xx_MOD_WAVEFORM_GAUSSIAN - Selects the gaussian noise waveform. */
' /*                                                                           */
' /*  PARAMETER SIX                                                            */
' /*    ViInt16 modPMIntRange                                                  */
' /*      Selects the phase modulation range based on the value of the phase   */
' /*      deviation.  The PM range is coupled to CW frequency, PM rate         */
' /*      (modPMIntFreqRate), and PM deviation (modPMIntDeviation).            */
' /*                                                                           */
' /*      hp837xx_MOD_PM_INT_RANGE_AUTO - Selects phase modulation range       */
' /*        automatically.  The range will automatically change from LOW to    */
' /*        HIGH or HIGH to LOW depending on which range will meet the new set */
' /*        of parameters.                                                     */
' /*      hp837xx_MOD_PM_INT_RANGE_LOW - Sets the phase modulation range from  */
' /*        15 mrads to 4 rads.                                                */
' /*      hp837xx_MOD_PM_INT_RANGE_HIGH - Sets the phase modulation range from */
' /*        .781 rads to 200 rads.                                             */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modPMConfigInt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modPMIntState As Integer, _
     ByVal modPMIntDeviation As Double, _
     ByVal modPMIntFreqRate As Double, _
     ByVal modPMIntWaveform As Integer, _
     ByVal modPMIntRange As Integer) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modPMConfigExt                                                   */
' /*    This method configures and sets the state of the external phase        */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    This method is only available on HP 83731B/32B models with Option 800  */
' /*    installed.                                                             */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modPMExtState                                                */
' /*      Sets phase modulation on or off.  When the synthesizer is set to the */
' /*      preset state, phase modulation is turned off.                        */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 modPMExtCoupling                                               */
' /*      Selects either AC or DC coupling for the FM/PM IN connector.         */
' /*                                                                           */
' /*      hp837xx_MOD_COUPLING_AC - Sets coupling to AC.                       */
' /*      hp837xx_MOD_COUPLING_DC - Sets coupling to DC                        */
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViReal64 modPMExtSensitivity                                           */
' /*      Sets the PM sensitivity in carrier deviation per volt.  Entered      */
' /*      values will automatically adjust to the closest preset value for the */
' /*      given CW frequency.  The allowable range for the parameter is shown  */
' /*      in the table below.  Note that there are only two values of PM       */
' /*      sensitivity available, dependent upon carrier frequency.             */
' /*                                                                           */
' /*      RANGE                   HIGH            LOW                          */
' /*      10 MHz to < 16 MHz      390 mrad/V      7.81 mrad/V                  */
' /*      16 MHz to < 64 MHz      781 mrad/V      15.6 mrad/V                  */
' /*      64 MHz to < 256 MHz     3.12 rad/V      62.5 mrad/V                  */
' /*      256 MHz to < 1 GHz      12.5 rad/V      250 mrad/V                   */
' /*      1 GHz to 20 GHz         50 rad/V        1 rad/V                      */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modPMConfigExt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modPMExtState As Integer, _
     ByVal modPMExtCoupling As Integer, _
     ByVal modPMExtSensitivity As Double) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modPulseConfigInt                                                */
' /*    This method configures and sets the state of the internal phase        */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    This method is only available for 83731B and 83732B.                   */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modPulseIntState                                             */
' /*      Turns pulse modulation on or off.  When the synthesizer is set to    */
' /*      the preset state, pulse modulation is turned off.                    */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 modPulseIntTrigMode                                            */
' /*      Selects the trigger source mode.                                     */
' /*                                                                           */
' /*      hp837xx_MOD_PULSE_TRIG_MODE_INTERNAL - Selects the immediate pulse   */
' /*        triggering.                                                        */
' /*      hp837xx_MOD_PULSE_TRIG_MODE_EXTERNAL - Sets the PULSE/TRIG GATE IN   */
' /*        connector as the trigger start source.  The trigger stop source is */
' /*        immediate (no external stop trigger).                              */
' /*      hp837xx_MOD_PULSE_TRIG_MODE_DOUBLET - Selects the pulse doublet      */
' /*        feature.  Each trigger at the PULSE/TRIG GATE IN connector will    */
' /*        produce two pulses.  The first pulse will follow the external      */
' /*        trigger signal.  The second pulse will have a delay and width as   */
' /*        set by the modPulseIntDelay and modPulseIntWidth parameters.       */
' /*      hp837xx_MOD_PULSE_TRIG_MODE_GATED - Sets the PULSE/TRIG GATE IN      */
' /*        connector as the trigger start source.  The trigger stop source is */
' /*        external                                                           */
' /*                                                                           */
' /*  PARAMETER FOUR                                                           */
' /*    ViReal64 modPulseIntDelay                                              */
' /*      Sets the pulse delay in ns.  The allowable range is -419 ms to +419  */
' /*      ms. The preset value for pulse delay is 1 us.                        */
' /*                                                                           */
' /*      Pulse delay entries have a resolution of 25 ns; entries with a       */
' /*      resolution finer than 25 ns will be rounded to the nearest 25 ns.    */
' /*                                                                           */
' /*  PARAMETER FIVE                                                           */
' /*    ViReal64 modPulseIntWidth                                              */
' /*      Sets the pulse width in ns.  The allowable range is 0 ns to 419 ms.  */
' /*      The preset value for pulse width is 10 us.                           */
' /*                                                                           */
' /*      Pulse width entries have a resolution of 25 ns; entries with a       */
' /*      resolution finer than 25 ns will be rounded to the nearest 25 ns.    */
' /*                                                                           */
' /*  PARAMETER SIX                                                            */
' /*    ViReal64 modPulseIntPeriod                                             */
' /*      Sets the pulse repetition interval in ns.  The allowable range is    */
' /*      300 ns to 419 ms.  The preset value for pulse repetition internal is */
' /*      100 us.                                                              */
' /*                                                                           */
' /*      The resolution is 25 ns; entries with a resolution finer than 25 ns  */
' /*      will be rounded to the nearest 25 ns.                                */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modPulseConfigInt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modPulseIntState As Integer, _
     ByVal modPulseIntTrigMode As Integer, _
     ByVal modPulseIntDelay As Double, _
     ByVal modPulseIntWidth As Double, _
     ByVal modPulseIntPeriod As Double) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_modPulseConfigExt                                                */
' /*    This method configures and sets the state of the external pulse        */
' /*    modulation.                                                            */
' /*                                                                           */
' /*    This method is only available for 83731B and 83732B.                   */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean modPulseExtState                                             */
' /*      Turns pulse modulation on or off.  When the synthesizer is set to    */
' /*      the preset state, pulse modulation is turned off.                    */
' /*                                                                           */
' /*  PARAMETER THREE                                                          */
' /*    ViInt16 modPulseExtPolarity                                            */
' /*      Selects either inverted or non-inverted polarity for the external    */
' /*      pulse input at the PULSE/TRIG GATE IN connector.                     */
' /*                                                                           */
' /*      hp837xx_MOD_PULSE_POL_NORMAL - Selects non-inverted polarity for the */
' /*                                     external pulse input.                 */
' /*      hp837xx_MOD_PULSE_POL_INVERTED - Selects inverted polarity for the   */
' /*                                       external pulse input.               */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_modPulseConfigExt Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal modPulseExtState As Integer, _
     ByVal modPulseExtPolarity As Integer) As Long
'
'
'         /*----------------------------------------------------------------*
'          |  Direct Functions                                              |
'          *----------------------------------------------------------------*/
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_frequencySetCW -                                                 */
' /*    This method sets the output frequency of the synthesizer.  The         */
' /*    frequency entered is the CW frequency if no modulation is chosen,      */
' /*    or the carrier frequency of any modulation type that is chosen.  The   */
' /*    preset value for the frequency parameter is 3 GHz.                     */
' /*                                                                           */
' /*    The allowable range for the frequency parameter is 1.0 GHz to 20GHz    */
' /*    for the HP 83711A/11B/31A/31B or 0.01 GHz to 20 Ghz for the            */
' /*    HP 83712A/12B/32A/32B.  Frequency resolution is 1 kHz.  If Option 1E8  */
' /*    is installed, frequency resolution is 1 Hz.                            */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViReal64 CWFrequency                                                   */
' /*      Sets the synthesizer output frequency.                               */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_frequencySetCW Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal CWFrequency As Double) As Long
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_frequencyGetCW -                                                 */
' /*    This method retrieves the current output frequency of the synthesizer. */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViPReal64 CWFrequency_Q                                                */
' /*      The synthesizer's current output frequency.                          */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_frequencyGetCW Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef CWFrequency_Q As Double) As Long
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_frequencySetStepSize -                                           */
' /*    This method selects the increment value for the synthesizer output     */
' /*    frequency.  When the hp837xx_frequencyStep method is used with the     */
' /*    "Up" or "Down" option, the output frequency will be increased or       */
' /*    decreased by the step size set with this method.                       */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViReal64 CWFrequencyStepSize                                           */
' /*      Sets the increment value for output frequency.  The allowable range  */
' /*      without option 1E8 for the parameter is 1 kHz to 19.99 GHz.  If      */
' /*      option 1E8 is installed, the allowable range for the parameter is 1  */
' /*      Hz to 19.99 GHz.                                                     */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_frequencySetStepSize Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal CWFrequencyStepSize As Double) As Long
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_frequencyStep -                                                  */
' /*    This method increases or decreases the synthesizer output frequency by */
' /*    the current output frequency increment value, as set by                */
' /*    hp837xx_frequencySetStepSize.                                          */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 CWFrequencyStepDirection                                       */
' /*      The direction to step the output frequency.                          */
' /*                                                                           */
' /*      hp837xx_FREQUENCY_STEP_UP - Steps the frequency up by the current    */
' /*                                    output frequency increment value.      */
' /*                                                                           */
' /*      hp837xx_FREQUENCY_STEP_DOWN - Steps the frequency down by the current*/
' /*                                    output frequency increment value.      */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_frequencyStep Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal CWFrequencyStepDirection As Integer) As Long
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_frequencySetMultiplier -                                         */
' /*    This method sets the multiplier value so that the synthesizer display  */
' /*    will indicate the frequency at the output of an external frequency     */
' /*    multiplier.                                                            */
' /*                                                                           */
' /*    Entering a frequency multiplier value is useful when an output         */
' /*    frequency will be generated with external multiplier equipment.        */
' /*    Setting the multiplier value scales the display so that the frequency  */
' /*    shown on the display will be the frequency at the output of the        */
' /*    external frequency multiplier, not at the synthesizer RF OUTPUT        */
' /*    connection.                                                            */
' /*                                                                           */
' /*    When the multiplier function is be used and you enter a frequency      */
' /*    parameter value with hp837xx_frequencySetCW, be aware that the entered */
' /*    frequency divided by the multiplier value (the frequency before        */
' /*    multiplication) has a minimum resolution of 1 kHz (1 Hz for Option     */
' /*    1E8).  As an example, assume a multiplier value of 2 has been entered  */
' /*    and you attempt to enter a frequency of 4,000,001,000 Hz.  The actual  */
' /*    frequency that the synthesizer would need to generate would be         */
' /*    2,000,000,500 Hz.  The synthesizer, however, can not output this       */
' /*    signal because the standard specified resolution is 1 kHz.  In this    */
' /*    case, the actual output frequency would be rounded to 2,000,001,000 Hz */
' /*    and the display would show 4,000,002,000 Hz.                           */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 CWFrequencyMultiplier                                          */
' /*      The new multiplier value.  The allowable range for the parameter is  */
' /*      1 to 100.  The preset value is 1.                                    */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_frequencySetMultiplier Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal CWFrequencyMultiplier As Integer) As Long
'
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_powerSetLevel -                                                  */
' /*    This method sets the output power level of the synthesizer.  Power     */
' /*    level resolution is 0.01 dB.                                           */
' /*                                                                           */
' /*    When the power level is modified, the synthesizer circuitry will       */
' /*    ensure that transitions from one power level to another will not allow */
' /*    the level to exceed the maximum of the two levels if the instrument is */
' /*    in CW mode (not modulated).  If the RF output is being amplitude       */
' /*    modulated or pulse modulated, the synthesizer circuitry will ensure    */
' /*    that transitions from one power level to another will not exceed the   */
' /*    maximum of the two power levels by more than 0.5 dB typically.         */
' /*                                                                           */
' /*    Changing the frequency or power level while pulse modulating the       */
' /*    output triggers an internal power level calibration.  This calibration */
' /*    includes a CW calibration for approximately 10 ms for HP 83731B/32B;   */
' /*    30 ms for HP 83731A/32A.  Refer to the hp832xx_powerSetProtectionState */
' /*    method for more information on how to protect devices sensitive to CW  */
' /*    power.                                                                 */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViReal64 powerLevel                                                    */
' /*      Sets the synthesizer output power level in dBm.  The allowable range */
' /*      is -120 dBm (-100 dBm on the HP 83711A/12A/31A/32A) to +30 dBm if    */
' /*      Option 1E1 is installed and -15 dBm  to +30 dBm if Option 1E1 is not */
' /*      installed.                                                           */
' /*                                                                           */
' /*      The preset value is -110 dBm (-90 dBm on HP 83711A/12A/31A/32A) if   */
' /*      option 1E1 is intalled and 0 dBm if Option 1E1 is not installed.     */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_powerSetLevel Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal powerLevel As Double) As Long
'
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_powerGetLevel -                                                  */
' /*    This method retrieves the current output power level of the            */
' /*    synthesizer.                                                           */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViPReal64 powerLevel_Q                                                 */
' /*      The synthesizer's current output power level in dBm.                 */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_powerGetLevel Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByRef powerLevel_Q As Double) As Long
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_powerSetStepSize -                                               */
' /*    This method selects the increment value for the synthesizer output     */
' /*    power level.  When "Up" or "Down" parameters are used with the         */
' /*    hp837xx_powerStep method, the output power level will be increased or  */
' /*    decreased by a step size set with this method.                         */
' /*                                                                           */
' /*    Numeric power level increment value entries have a resolution of       */
' /*    0.01 dB.                                                               */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViReal64 powerLevelStepSize                                            */
' /*      Sets the increment value for output power level in dB.  The          */
' /*      allowable range is 0.01 dB to 150 dB if Option 1E1 is installed and  */
' /*      0.01 dB to 45 dB if option 1E1 is not installed.                     */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_powerSetStepSize Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal powerLevelStepSize As Double) As Long
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_powerStep -                                                      */
' /*    This method increases or decreases the synthesizer's output power      */
' /*    level by the current power level increment value.                      */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 powerStepDirection                                             */
' /*      The direction to step the synthesizer's output power.                */
' /*                                                                           */
' /*      hp837xx_POWER_STEP_UP - Steps the power level up by the current      */
' /*      output power level increment value.                                  */
' /*                                                                           */
' /*      hp837xx_POWER_STEP_DOWN - Steps the power level down by the          */
' /*      current output power level increment value.                          */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_powerStep Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal powerStepDirection As Integer) As Long
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_powerSetState -                                                  */
' /*    This method turns the signal at the RF OUTPUT connector on or off.     */
' /*                                                                           */
' /*    When setting the state to off, the internal oscillators are turned off */
' /*    and the internal RF power shutdown circuit is turned on.               */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean powerStateOn                                                 */
' /*      The state of the signal at the RF OUTPUT connector.  The preset      */
' /*      value is on.                                                         */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_powerSetState Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal powerState As Integer) As Long
'
'
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_powerSetProtectionState -                                        */
' /*    This method turns the average power inhibit function on or off.  This  */
' /*    is not available with HP 83711/12 instrument or if option 1E1 (step    */
' /*    attenuator) is not installed.                                          */
' /*                                                                           */
' /*    The average power inhibit function can be used during pulse modulation */
' /*    to protect devices sensitive to high average power.  When the output   */
' /*    power level or frequency of the synthesizer is changed during pulse    */
' /*    modulation, the internal leveling algorithm causes the RF output to be */
' /*    momentarily switched to CW to enable the synthesizer circuitry to      */
' /*    sample the signal level and make a correction.  If the output of the   */
' /*    synthesizer is connected to circuitry that is average power-sensitive, */
' /*    damage to the circuitry could result during this CW calibration.  When */
' /*    in internal leveling mode, the CW calibration is approximately 30 ms.  */
' /*                                                                           */
' /*    When the average power inhibit function is off (the preset condition), */
' /*    the CW calibration will accompany output power level and frequency     */
' /*    changes.  The CW calibration will also be present the first time pulse */
' /*    or logarithmic amplitude modulation is enabled.  When average power    */
' /*    inhibit is on, the internal step attenuator will switch to 110 dB (90  */
' /*    dB on HP 83731A/32A) of attenuation during the CW calibration.  This   */
' /*    will protect power-sensitive circuitry connected to the RF OUTPUT      */
' /*    connector, but will cause extra wear on the step attenuator.  Turning  */
' /*    the function on will also cause a momentary drop in signal power       */
' /*    (approximately 200 ms) and will lengthen frequency and power level     */
' /*    switching times by 70 ms.                                              */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean powerProtectionStateOn                                       */
' /*      The state of the average power inhibit function.  When the           */
' /*      synthesizer is set to the preset state, the average power inhibit    */
' /*      function is turned off.                                              */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_powerSetProtectionState Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal powerProtectionState As Integer) As Long
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_outputSetProtectionState                                         */
' /*    This method turns RF protection during frequency switching on or off.  */
' /*    This is useful when measuring the synthesizer frequency switching time.*/
' /*                                                                           */
' /*    The synthesizer contains an RF protection circuit that momentarily     */
' /*    attenuates output power and then brings the output power back up to    */
' /*    the required level (in 20 ms nominal) when the synthesizer output      */
' /*    frequency is changed.  This circuit assures that the output power does */
' /*    not overshoot the power level set via the front panel or HP-IB during  */
' /*    frequency switching.                                                   */
' /*                                                                           */
' /*    RF protection during frequency switching can not be turned off when    */
' /*    AM, FM, or pulse modulation is being used.  It can only be turned off  */
' /*    when the synthesizer is in CW mode.                                    */
' /*                                                                           */
' /*    Even when the synthesizer is in CW mode and the RF protection during   */
' /*    frequency switching function is turned off, the RF protection circuit  */
' /*    will switch in when the synthesizer divider circuits switch or         */
' /*    whenever frequency switches greater than 260 MHz occur.                */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViBoolean outputProtectionStateOn                                      */
' /*      The new state of the RF protection during frequency switching.  When */
' /*      the synthesizer is set to the preset state, RF protection is turned  */
' /*      on.                                                                  */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_outputSetProtectionState Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal outputProtectionState As Integer) As Long
'
'
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_outputSetAttnHoldState -                                         */
' /*    This method turns the attenuator hold function on or off.  The         */
' /*    attenuator hold function can be used to extend the vernier range to    */
' /*    prevent the step attenuator from switching between two levels.         */
' /*    Locking the step attenuator keeps the attenuator from switching        */
' /*    between two levels as leveled power is varied above and below the      */
' /*    threshold level, thus saving wear on the attenuator.                   */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 outputAttnHoldState                                            */
' /*      The new state of the attenuator hold function.                       */
' /*                                                                           */
' /*      hp837xx_ATTENUATOR_HOLD_ON - Turns on the attenuator hold function.  */
' /*                                                                           */
' /*      hp837xx_ATTENUATOR_HOLD_OFF - Turns off the attenuator hold function.*/
' /*                                                                           */
' /*      hp837xx_ATTENUATOR_HOLD_ONCE - Temporarily turns the attenuator hold */
' /*      function off so that the synthesizer can automatically update the    */
' /*      attenuator setting.  It is then turned on to lock the attenuator at  */
' /*      that setting.                                                        */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_outputSetAttnHoldState Lib "hp837xx_32.dll" _
        (ByVal instrumentHandle As Long, _
         ByVal outputAttnHoldState As Integer) As Long
'
'
' /****************************/
' /* Memory Control Functions */
' /****************************/
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_instrmentStateSave                                               */
' /*    This method saves the instrument state in one of ten register          */
' /*    locations.  All user settings that are affected by preset will be      */
' /*    saved.  Level correction tables will not be saved.                     */
' /*                                                                           */
' /*    Saving the instrument state to a given register location will write    */
' /*    over any instrument state previously stored in that register.          */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 saveRegister                                                   */
' /*      The number of the register where the instrument state is to be       */
' /*      stored.  The number must be an integer from 0 to 9.                  */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_instrumentStateSave Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal saveRegister As Integer) As Long
'
'
'
' /*****************************************************************************/
' /*                                                                           */
' /*  hp837xx_instrumentStateRecall                                            */
' /*    This method recalls a previously stored instrument state from one of   */
' /*    ten register locations.                                                */
' /*                                                                           */
' /*    If the register location has not been previous saved, the preset state */
' /*    is recalled.                                                           */
' /*                                                                           */
' /*  PARAMETER ONE                                                            */
' /*    ViSession instrumentHandle                                             */
' /*      Instrument handle returned from hp837xx_init().                      */
' /*                                                                           */
' /*  PARAMETER TWO                                                            */
' /*    ViInt16 recallRegister                                                 */
' /*      The number of the register where the desired instrument state has    */
' /*      been stored.  The number must be an integer from 0 to 9.             */
' /*                                                                           */
' /*---------------------------------------------------------------------------*/
Declare Function hp837xx_instrumentStateRecall Lib "hp837xx_32.dll" _
    (ByVal instrumentHandle As Long, _
     ByVal recallRegister As Integer) As Long
'
'
'
'
' /*---------------------------------------------------------------------------*/
' /*                        end Function Prototypes                            */
' /*                                                                           */
' /*****************************************************************************/
