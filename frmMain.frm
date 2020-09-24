VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GemEdit by Disk2 - [No File]"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   1650
      Top             =   975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   32000
   End
   Begin VB.ComboBox cmbSource 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   75
      List            =   "frmMain.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   300
      Width           =   1440
   End
   Begin VB.ComboBox cmbDest 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":046E
      Left            =   75
      List            =   "frmMain.frx":0487
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1275
      Width           =   1440
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Character"
      Height          =   390
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   75
      Width           =   1965
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert Potions to Gems"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   525
      Width           =   1965
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   390
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1425
      Width           =   1965
   End
   Begin VB.CommandButton cmdUpgrade 
      Caption         =   "&Upgrade All To Perfect"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   975
      Width           =   1965
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   390
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1875
      Width           =   1965
   End
   Begin VB.Label lblInfo 
      Caption         =   "Source Type:"
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   1440
   End
   Begin VB.Label lblInfo 
      Caption         =   "Destination Type:"
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   7
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   2190
      Left            =   1650
      Picture         =   "frmMain.frx":04CD
      Stretch         =   -1  'True
      Top             =   75
      Width           =   30
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GEMEDIT v1.10 SOURCE CODE                                                       '
' Written by Disk2 (disktwo@yahoo.com)                                            '
'                                                                                 '
' This is the source code to the GemEdit program I made for Diablo 2.             '
' It's nothing fancy, but it works. I've tried to comment the code fairly well.   '
' If you have any questions, PLEASE don't email me :) Figure it out for yourself. '
' I don't have the time to help, and I think you learn better by trying.          '
'                                                                                 '
' I released this code to help anyone who wants to make an editor. I shows        '
' basic binary I/O. It's a start :)                                               '
'                                                                                 '
' If you use this code, I request a mention somewhere in the editor :) Thanks...  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Used to find the beginning and end of the inventory section...
Private Type ItemHeader
    szFirstJM As String * 2
    iItemCount As Byte
    iEmpty As Byte
    szLastJM As String * 2
End Type

' This isn't complete :) It does what I need it to though.
Private Type Item
    iSubType As Byte
    iType As Byte
End Type

' Holds the filename and declares the item header...
Dim strFileName As String
Dim ItemHead As ItemHeader

Private Sub cmdAbout_Click()
    ' Show the about info
    MsgBox "GemEdit v1.10" & vbCrLf & vbCrLf & "Written by Disk2" & vbCrLf & vbCrLf & "This program should work in Diablo 2 v1.00, 1.01, and 1.02." & vbCrLf & "If it doesn't, please tell email me at disktwo@yahoo.com.", vbOKOnly + vbInformationm, "About GemEdit"
End Sub

Private Sub cmdExit_Click()
    ' Exit
    End
End Sub

Private Sub cmdConvert_Click()
    ' Check to see if the user select a source and destination type...
    If cmbSource.Text = "" Or cmbDest.Text = "" Then
        MsgBox "You must select a source and destination type before converting!", vbCritical + vbOKOnly, "Error"
    Else
        ' If the user DID, then convert the items.
        Convert
    End If
End Sub

Private Sub cmdOpen_Click()
    ' Set the common dialog properties
    dlgMain.Flags = &H1000
    dlgMain.InitDir = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\Diablo II\", "Save Path")
    dlgMain.DialogTitle = "Open Character"
    dlgMain.Filter = "Diablo 2 Saved Games (*.d2s)|*.d2s|"
    dlgMain.CancelError = False
    dlgMain.MaxFileSize = 32000
    dlgMain.ShowOpen
    
    ' If the user selected a file, open it and enable the controls...
    If Len(dlgMain.FileName) > 0 Then
        strFileName = dlgMain.FileName
        
        frmMain.Caption = "GemEdit by Disk2 - [" & dlgMain.FileTitle & "]"
        
        cmbSource.Enabled = True
        cmbDest.Enabled = True
        cmdConvert.Enabled = True
        cmdUpgrade.Enabled = True
    End If
End Sub

Private Sub cmdUpgrade_Click()
    ' Fix the user's gems... This is done in the FixGems sub
    FixGems
End Sub

Private Sub FixGems()
    On Error Resume Next
    
    Dim iPos As Integer ' Holds the position in the file
    Dim xItem As Item ' Temp item
    Dim TheString As String * 4 ' will be JMJM if we're at the end of the items...
    Dim TheEnd As ItemHeader ' Used to find the end of the items

    ' Reset ItemHead
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    iPos = &H1 ' Start at the beginning of the file

    ' Open the file
    Open strFileName For Binary As #1
        ' Read from the file until we find the "JM  JM". This means we've found the beginning of the item data...
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1 ' Increase the position
        Loop
    
        iPos = iPos + 3 ' Go to the REAL start of the item information.

        ' If the user has no items, don't continue...
        If ItemHead.iItemCount = 0 Then
            MsgBox "This character doesn't appear to have any items! If this is an error please email me at cregistry@yahoo.com and attach the saved game file. Thanks!", vbOKOnly + vbInformation, "Notice"
            
            Close #1 ' Close the file
            Exit Sub
        End If

        ' Read items. Compare them with known gem codes. Convert them if they aren't perfect.
        Do Until TheString = "JMJM"
            Get #1, iPos, TheEnd
            
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            iPos = iPos + 2
        
            Get #1, iPos + 6, xItem.iSubType
            Get #1, iPos + 7, xItem.iType
            
            ' Hehe. This is confusing.
            ' The actual concept isn't. It's just the way I coded it :)
            ' Here's a table of all the gem codes...
            '
            '          | Chipped | Flawed | Regular | Flawless | Perfect |
            ' ---------|--------------------------------------------------
            ' Diamond  | 5015    | 6015   | 7015    | 8015     | 9015    |
            ' ---------|--------------------------------------------------
            ' Ruby     | 1015    | 0015   | 2015    | 3015     | 4015    |
            ' ---------|--------------------------------------------------
            ' Topaz    | 1014    | 2014   | 3014    | 4014     | 5014    |
            ' ---------|--------------------------------------------------
            ' Sapphire | 6014    | 7014   | 8014    | 9014     | A014    |
            ' ---------|--------------------------------------------------
            ' Amethyst | C013    | D013   | E013    | F013     | 0014    |
            ' ---------|--------------------------------------------------
            ' Emerald  | B014    | C014   | D014    | E014     | F014    |
            ' ---------|--------------------------------------------------
            ' Skull    | 4016    | 5016   | 6016    | 7016     | 8016    |
            ' ---------|--------------------------------------------------
            
            ' With that in mind, you can figure out this code.
            Select Case xItem.iType
                Case &H13
                    Select Case xItem.iSubType
                        Case &HD0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HC0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HF0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HE0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                    End Select
                Case &H14
                    Select Case xItem.iSubType
                        Case &H20
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H10
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H40
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H30
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H60
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H70
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H80
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H90
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &HC0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HB0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HE0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HD0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                    End Select
                    Case &H15
                    Select Case xItem.iSubType
                        Case &H60
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H50
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H80
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H70
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H0
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H10
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H20
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H30
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                    End Select
                    Case &H16
                    Select Case xItem.iSubType
                        Case &H40
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H50
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H60
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H70
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                    End Select
            End Select
            
            ' Increase the position so we can read the next item.
            iPos = iPos + 25
        Loop
    Close #1
    
    ' Tell the user the gems were perfected
    MsgBox "All your gems are now perfect.", vbOKOnly + vbInformation, "Done"
End Sub

Private Sub Convert()
    On Error Resume Next

    Dim iPos As Integer
    Dim xItem As Item
    Dim dItem As Item
    Dim TheEnd As ItemHeader
    Dim TheString As String * 4
    
    ' Depending of the destination type, set the temp item type to a perfect gem.
    ' Refer to the table in FixGems for the gem codes...
    Select Case cmbDest.Text
        Case "Diamonds"
            dItem.iType = &H15
            dItem.iSubType = &H90
        Case "Rubys"
            dItem.iType = &H15
            dItem.iSubType = &H40
        Case "Topazes"
            dItem.iType = &H14
            dItem.iSubType = &H50
        Case "Sapphires"
            dItem.iType = &H14
            dItem.iSubType = &HA0
        Case "Amethysts"
            dItem.iType = &H14
            dItem.iSubType = &H0
        Case "Emeralds"
            dItem.iType = &H14
            dItem.iSubType = &HF0
        Case "Skulls"
            dItem.iType = &H16
            dItem.iSubType = &H80
    End Select

    ' Reset ItemHead
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    ' Start at the beginning of the file
    iPos = &H1

    ' Open the file
    Open strFileName For Binary As #1
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1
        Loop
    
         ' Go to the REAL item data start
        iPos = iPos + 3

        ' If the user has no items, there's no point in continuing.
        If ItemHead.iItemCount = 0 Then
            MsgBox "This character doesn't appear to have any items! If this is an error please email me at cregistry@yahoo.com and attach the saved game file. Thanks!", vbOKOnly + vbInformation, "Notice"
            
            Close #1
            Exit Sub
        End If

        ' Read items until we reach the end of the file.
        Do Until TheString = "JMJM"
            Get #1, iPos, TheEnd
            
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            iPos = iPos + 2
            
            ' Get the item type
            Get #1, iPos + 6, xItem.iSubType
            Get #1, iPos + 7, xItem.iType
            
            ' Depending on the source type, look for health or mana potions.
            ' You can figure out the potion codes by looking at the code below.
            If cmbSource.Text = "Health Potions" Then
                If xItem.iType = &H15 Then
                    If xItem.iSubType = &HA0 Or xItem.iSubType = &HB0 Or xItem.iSubType = &HC0 Or xItem.iSubType = &HD0 Or xItem.iSubType = &HE0 Then
                        Put #1, iPos + 6, dItem
                    End If
                End If
            End If
            If cmbSource.Text = "Mana Potions" Then
                If xItem.iType = &H16 Then
                    If xItem.iSubType = &H0 Or xItem.iSubType = &H10 Or xItem.iSubType = &H20 Or xItem.iSubType = &H30 Then
                        Put #1, iPos + 6, dItem
                    End If
                ElseIf xItem.iType = &H15 Then
                    If xItem.iSubType = &HF0 Then
                        Put #1, iPos + 6, dItem
                    End If
                End If
            End If
            
            ' Increase the position so we can read the next file
            iPos = iPos + 25
        Loop
    Close #1

    ' Show the message...
    MsgBox "All your " & cmbSource.Text & " are now " & cmbDest.Text & ".", vbOKOnly + vbInformation, "Success"
End Sub

