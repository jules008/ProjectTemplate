Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Private Const StrMODULE As String = "ModGlobals"
Public Const DEBUG_MODE As Boolean = True   ' TRUE / FALSE
Public Const OUTPUT_MODE As String = "Debug"  ' "Log" / "Debug"
Public Const ENABLE_PRINT = False           ' TRUE / FALSE
Public Const APP_NAME As String = "Phase 2 Database"
Public Const HANDLED_ERROR As Long = 9999
Public Const USER_CANCEL As Long = 18
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const VERSION = "0.0"
Public Const VER_DATE = "11/01/17"

' ===============================================================
' Global Variables
' ---------------------------------------------------------------
Public DB_PATH As String

' ===============================================================
' Global Class Declarations
' ---------------------------------------------------------------
Public MailSystem As ClsMailSystem

