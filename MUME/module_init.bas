Attribute VB_Name = "initVariables"
Option Explicit
Public tsehhisalasona As String
Public Const mapTitle = "C:\MOM.."
Public Const DEBUGMODE = False
Public Const GODMODE = True

Public Sub loadVariables()
   MUDname = "MUME"
   virtualRow = 32
   virtualCol = 250
   theROW = virtualRow
   theCOL = virtualCol
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
   indexEnemies = 0
'set mom defaults
   frmMap.mnuWalkthrough.Checked = False
   frmMap.mnuEdit.Enabled = True ' wasfalse
   frmMap.mnuEdit.Visible = True ' wasfalse
   frmMap.mnuRoomsync.Checked = False
   frmMap.mnuGroup.Checked = False
   frmMap.mnuAutosync.Checked = True
   frmMap.mnuPortals.Checked = True         'will be true
   frmMap.mnuAlwaysOnTop.Checked = True      'will be true
   frmMap.mnuNotes.Checked = True            'will be true
   frmMap.mnuAutosync.Checked = True         'will be true
   frmMap.mnuBrief.Checked = False           'will be true
   frmMap.mnuSpam.Checked = True             'will be false
   'frmMap.mnuFollow.Checked = False: followMode = False
   frmMap.mnuFollow.Checked = False: frmMap.mnuFollow.Enabled = True: frmMap.mnuFollow.Visible = True

'load user defaults
   Call loadMOMini
   
   If Not GODMODE Then
      frmMap.mnuMovement.Checked = False: frmMap.mnuMovement.Enabled = False: frmMap.mnuMovement.Visible = False
      frmMap.mnuPlayers.Checked = False: frmMap.mnuPlayers.Enabled = False: frmMap.mnuPlayers.Visible = False
      frmMap.mnuEnemies.Checked = False: frmMap.mnuEnemies.Enabled = False: frmMap.mnuEnemies.Visible = False
      frmMap.mnuTarget.Checked = False: frmMap.mnuTarget.Enabled = False: frmMap.mnuTarget.Visible = False
      frmMap.mnuHere.Checked = False: frmMap.mnuHere.Enabled = False: frmMap.mnuHere.Visible = False
      frmMap.mnuWalk.Checked = False: frmMap.mnuWalk.Enabled = False: frmMap.mnuWalk.Visible = False
      frmMap.mnuReceiver.Checked = False: frmMap.mnuReceiver.Enabled = False: frmMap.mnuReceiver.Visible = False
      frmMap.mnuInformer.Checked = False: frmMap.mnuInformer.Enabled = False: frmMap.mnuInformer.Visible = False
   End If
   
   Call frmTools.Sun_Click
   Call frmTools.Ridable_Click
   frmMap.Caption = mapTitle
End Sub
Public Function getPassword()
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim systemID As String, didder As Boolean
   didder = False
   systemID = fso.GetDrive(Mid(systemRoot, 1, 2)).SerialNumber

   If GODMODE Then ' check if god exists
      Select Case systemID
      Case "-1875040727" 'jaanuslang tööarvuti
        didder = True
      Case "3531849" 'kaspar metsa
        didder = True
      Case "442741118" 'naga
        didder = True
      Case "-2133279939"
        didder = True 'jyri abramov.. uus arvuti
      Case "-257228252" 'horus as fredriktjust@hotmail.com
         didder = True
      Case "-1127093233"
         didder = True 'priit kahn
      Case "1151102848", "137263138"
         didder = True 'kalm
      Case "1280610282"
         didder = True 'alan kesselman
      Case "-532512855"
         didder = True 'marduk
      Case "1881301932"
         didder = True 'hardi
      Case "-1736397739"
         didder = True 'priit kahn
      Case "-62492529"
         didder = True 'rainer, timbulimbu
      Case "-598341821" 'alanke kodus
         didder = True
      Case "1625896582" 'ivar songe/godzilla
         didder = True
      Case "526456556" 'alanke, tööarvuti
         didder = True 'ypsilon
      Case "-1810568471" 'michael prill, deor, mikala, shorty
         didder = True 'ypsilon
      Case "-399713147" 'christen-johansen@hotmail.com
         didder = True 'ypsilon
      Case "415114316"
         didder = True 'ypsilon
      Case "1491964172"
         didder = True 'robert.. sloeveniast
      Case "1012508241"
         didder = True 'mattias liivak laptop
      Case "220337166"
         didder = True 'mattias liivak
      Case "-2143011068"
         didder = True 'nimrod topkin, mume njuubi.. mängis kunagi stonias
      Case "-598341821"
         didder = True 'alan kesselman/hector, kodus
      Case "1025773800"
         didder = True 'alan kesselman/hector
      Case "526456556"
         didder = True 'alan kesselman/hector
      Case "1756842553"
         didder = True ' marduk
      Case "-1130134776"
         didder = True ' gryyn
      Case "1814472136"
         didder = True ' gryyni vend
      Case "685095596"
         didder = True ' andre karpitsenko
      Case "1557838895"
         didder = True ' kadi niitenberg, kerti sõbrants
      Case "1814129021"
         didder = True ' priit padar, yerba
      Case "1891608374"
         didder = True ' priit padar, yerba, tööl
      Case "-1527422127"
         didder = True ' svenvaldmann uus töömasin
      Case "1420136176" 'fero, rehakcz(msn)
         didder = True
      Case "1505079639"
         didder = True
      Case "-1272747669"
         didder = True
      Case "-2072735926" 'Kaido Haavandi/stormblast
         didder = True
      Case "-532865964" 'arved järvet, sveni labi uus masin
         didder = True
      Case "-1129690231" 'Kaido Haavandi/stormblast
         didder = True
      Case "-58147556" 'kerti alev laptop
        didder = True
      Case "-721882785" 'kerti alev kodus #2
        didder = True
      Case "-2003596246"   'kerti alev tööl
        didder = True
      Case "1277289804" 'kerti alev kodus #1
        didder = True
      Case "-1875550103" ' andres mägi, aryan
        didder = True
      Case "-331665036" ' andres murdvee, murka
        didder = True
      Case "204004893" ' Björn Wretfeldt
        didder = True
      Case "-1943403421"   ' uus ivo lillepea lapakas
         didder = True
      Case "1485720930" 'ares hubel uus
         didder = True
      Case "1007088660" 'mammoth uus, jyri abramov
         didder = True
      Case "-1128999751" 'kristjan/hambaork, ' vana "-1128999751"
         didder = True
      Case "-1469130607" 'ivo lillepea/kozlor/donald
         didder = True
      Case "-1399096745" 'viljar vaht/gilgalad, uus arvuti
         didder = True
      Case "945241773"
         didder = True
      Case "-128391651"  'jyri etverk/focus
         didder = True
      Case "1880378297" 'uus tööarvuti
         didder = True
      Case "-2009860760" 'rait/maxam
         didder = True
      Case "145949111" 'jonas averling
         didder = True
      Case "685310091" 'liina pehka
         didder = True
      Case "272440080" 'ert pehka
         didder = True
      Case "-600235967" 'arved kodus
         didder = True
      Case "1142177182" 'sven kodu
         didder = True
      Case "-1068419196" 'uus tööarvuti
         didder = True
      Case "-737152708" 'harri teder, urkar evil, naksur
         didder = True
      Case "-1203671468" 'taavi vallner/tsort kodus
         didder = True
      Case "1154956820" 'taavi vallner tööl/tsort tööl
         didder = True
      Case "876287986"  'andres mägi
         didder = True
      Case "1143805954"    'mägi andres
         didder = True
      Case "-393936917"     'kodu c ketas!
         didder = True
      Case "1153790082"    'toomas valdmann-i masin
         didder = True
      Case "-399987758"    'svenerik valdmann
         didder = True
      Case "-1537604571" 'svenerik valdmann
         didder = True
      Case "-1332932192"   'andre karpitsenko pc
         didder = True
      Case "943601828" 'andre karpitsenko laptop
         didder = True
      Case "-1057348976"   'jyri abramov
         didder = True
      Case "2019800328" 'vikato
         didder = True
      Case "1128975847" 'marduk
         didder = True
      Case "-1675281392"    'marduk /2023764704
         didder = True
      Case "-125132484"    'it grupp tööl
         didder = True
      Case "-1461142044" 'Karli Grynberg
         didder = True
      Case "-1943024925" 'Kristo Grynberg
         didder = True
      Case "-1138900935" 'karli grynberg tööl
         didder = True
      Case "-51955447"  'tarmo tubro
         didder = True
      Case "-1802443727" 'taavi padjus
        didder = True
      Case "-664327385" 'alan kesselman
        didder = True
      Case "-1675797285" 'arved
        didder = True
      Case "1142661516" 'ivo lillpea
        didder = True
      Case "3531849" ' kassu17777... keegi kaspar
        didder = True
      Case "1793765703"
        didder = True 'lukas navratil
    Case "5861261" ' deor
        didder = True
    Case "-1875040727" '2in uus t55arvuti
        didder = True
    Case "1751819968" 'lukas 2, krvavjej
        didder = True
    Case "-1530391047" 'priit kahn / staub
        didder = True
    Case "-265363738"
        didder = True
    Case "138545486" 'vikat
        didder = True
    Case "-1933510598" 'ares
        didder = True
    Case "-601873875" 'kozlor lapakas
        didder = True
    Case "1546561322" 'luke kodus
        didder = True
    Case "-1460052452" 'joosep
        didder = True
    Case "-58147556" 'kerti alev
        didder = True
    Case "1010765173" 'Reef / Al Sutherland
        didder = True
    Case "1752042713" 'bjorn
        didder = True
    Case Else
        MsgBox "Invalid installation or corrupted database!" & vbCrLf & vbCrLf
        End
    End Select
   End If
   If didder Then systemID = "-125132484"
   getPassword = systemID

Exit Function
'original = cast128.cast128decode(key, encrypted)
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "fetch"
   writeError (errorModule)
End Function
