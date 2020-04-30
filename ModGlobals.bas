Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 28 Apr 20
'===============================================================
Private Const StrMODULE As String = "ModGlobals"

Option Explicit

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Public Const PROJECT_FILE_NAME As String = "CDC Tracker"
Public Const APP_NAME As String = "CDC Tracker"
Public Const EXPORT_FILE_PATH As String = "G:\Development Areas\CDC Tracker\Library\"
Public Const IMPORT_FILE_PATH As String = "G:\CDCTracker\"
Public Const INI_FILE_PATH As String = "\System Files\"
Public Const INI_FILE_NAME As String = "System.ini"
Public Const PROTECT_ON As Boolean = True
Public Const STOP_FLAG As Boolean = False
Public Const MAINT_MSG As String = ""
Public Const SEND_ERR_MSG As Boolean = False
Public Const TEST_PREFIX As String = "TEST - "
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const VERSION = "V0.0.0"
Public Const DB_VER = "V0.0.0"
Public Const VER_DATE = ""

' ===============================================================
' Error Constants
' ---------------------------------------------------------------
Public Const HANDLED_ERROR As Long = 9999
Public Const UNKNOWN_USER As Long = 1000
Public Const SYSTEM_RESTART As Long = 1001
Public Const NO_DATABASE_FOUND As Long = 1002
Public Const ACCESS_DENIED As Long = 1003
Public Const NO_INI_FILE As Long = 1004
Public Const DB_WRONG_VER As Long = 1005
Public Const GENERIC_ERROR As Long = 1006
Public Const FORM_INPUT_EMPTY As Long = 1007
Public Const NO_USER_SELECTED As Long = 1008

' ===============================================================
' Error Variables
' ---------------------------------------------------------------
Public FaultCount1002 As Integer
Public FaultCount1008 As Integer

' ===============================================================
' Global Variables
' ---------------------------------------------------------------
Public DEBUG_MODE As Boolean
Public SEND_EMAILS As Boolean
Public ENABLE_PRINT As Boolean
Public DB_PATH As String
Public DB_FILE_NAME As String
Public DEV_MODE As Boolean
Public SYS_PATH As String
Public CURRENT_USER As String
Public MENU_ITEM_SEL As Integer

' ===============================================================
' Global UI Class Declarations
' ---------------------------------------------------------------
Public MainScreen As ClsUIScreen
Public MailSystem As ClsMailSystem
Public CurrentUser As ClsMember
Public MenuBar As ClsUIFrame
Public Logo As ClsUIDashObj
Public Menu As ClsUIMenu
Public MenuItem As ClsUIMenuItem
Public MainFrame As ClsUIFrame
Public LeftFrame As ClsUIFrame
Public RightFrame As ClsUIFrame
Public Header As ClsUIHeader
Public BtnNewWorkflow As ClsUIMenuItem

' ===============================================================
' Global UI Class Declarations
' ---------------------------------------------------------------
Public MailSystem As ClsMailSystem
Public CurrentUser As ClsMember
Public Workflows As ClsWorkflows
' ---------------------------------------------------------------
' Others
' ---------------------------------------------------------------

' ===============================================================
' Colours
' ---------------------------------------------------------------
Public Const COLOUR_1 As Long = 12298644
Public Const COLOUR_2 As Long = 10185799
Public Const COLOUR_3 As Long = 9148699
Public Const COLOUR_4 As Long = 1383935
Public Const COLOUR_5 As Long = 4670781
Public Const COLOUR_6 As Long = 3450623
Public Const COLOUR_7 As Long = 6893787
Public Const COLOUR_8 As Long = 16056312
Public Const COLOUR_9 As Long = 12439241
Public Const COLOUR_10 As Long = 7864234
Public Const COLOUR_11 As Long = 52479

' ===============================================================
' Type Declarations
' ---------------------------------------------------------------
Type TypeStyle
    ForeColour As Long
    BorderColour As Long
    BorderWidth As Long
    FontStyle As String
    FontBold As Boolean
    FontSize As Integer
    FontColour As Long
    FontXJust As XlHAlign
    FontYJust As XlVAlign
    Fill1 As Long
    Fill2 As Long
    Shadow As MsoShadowType
End Type

' ===============================================================
' Enum Declarations
' ---------------------------------------------------------------
Enum EnumTriState
    xTrue
    xFalse
    xError
End Enum

Enum EnumObjType
    ObjImage = 1
    ObjChart = 2
End Enum

Enum EnumBtnNo
    EnumMyStation = 1
    EnumStores
    EnumReports
    EnumMyProfile
    EnumSupport
    EnumExit
    EnumNewOrder
    EnumSupportMsg
    EnumUserMngt
    EnumOrderSwitch
    EnumRemoteOrder
    EnumSupplierBtn
    EnumReport1Btn
    EnumManageDataBtn
    EnumReport2Btn
    EnumFindOrderBtn
    EnumReport3Btn
    EnumRptSettings
    EnumReturnStock
End Enum

' ===============================================================
' Enum Return Functions
' ---------------------------------------------------------------
Public Function EnRetStepStatus(EnumValue As EnStepStatus) As String
    Select Case EnumValue
        Case enNotStarted
            EnRetStepStatus = "Not Started"
        Case enStatGreen
            EnRetStepStatus = "In Progress - Green"
        Case enStatAmber
            EnRetStepStatus = "In progress - Amber"
        Case enStatRed
            EnRetStepStatus = "In progress - Red"
        Case enWait
            EnRetStepStatus = "Waiting"
        Case enComplete
            EnRetStepStatus = "Complete"
    End Select
End Function


