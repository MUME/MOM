Attribute VB_Name = "initVariables"
Option Explicit
Public Const DEBUGMODE = False

Public Sub loadVariables()
   virtualRow = 32
   virtualCol = 250
   theRow = virtualRow
   theCol = virtualCol
   roomcount = 0
   MappingData = False
   MappingGetUpdate = False
   dataFromMUD = False
   surfing = False
   wasMapMode = False
   canUndo = False
   LOST = True
   tmpOutput = ""
   limit = 3
   fleeRetry = 4
   selectType = 0

'set mom defaults
   'frmMap.mnuTools.Checked = False
   frmMap.mnuFollow.Checked = False: followMode = False
   frmMap.mnuEdit.Enabled = False
   frmMap.mnuRoomsync.Checked = False
   frmMap.mnuGroup.Checked = False
   frmMap.mnuAutosync.Checked = False
   frmMap.mnuBrief.Checked = False           'will be true
   frmMap.mnuSpam.Checked = True             'will be false
   frmMap.mnuAlwaysOnTop.Checked = False     'will be true
'load user defaults
   Call loadMOMini
   
   Call frmTools.Sun_Click
   Call frmTools.Ridable_Click
   frmMap.Caption = "MUD Online Map"
End Sub

Public Function getPassword()

If DEBUGMODE = False Then On Error GoTo errorhandler
   Dim systemID As String
   systemID = fso.GetDrive(Mid(systemRoot, 1, 2)).SerialNumber
   
'   Select Case systemID
'   Case "351659636"              'gosak new laptop key
'      systemID = "-325362215"    'his pc key, where the map is made
'   End Select
   
   getPassword = systemID
Exit Function

errorhandler:
   errorModule = Err.Description & "(" & Err.Number & ") -> " & "fetch"
   writeError (errorModule)
End Function
