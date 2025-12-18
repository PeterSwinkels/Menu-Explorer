VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MenuExplorerWindow 
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6570
   ClipControls    =   0   'False
   Icon            =   "Menu Explorer.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   18.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   54.75
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox MenuAncestorsBox 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "The active menu's parent menus."
      Top             =   3840
      Width           =   6255
   End
   Begin MSFlexGridLib.MSFlexGrid MenuWindowListBox 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "The list of windows that contain menus."
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid MenuListBox 
      Height          =   3615
      Left            =   3360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "The list of menus in the active menu."
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu ProgramSeparator1Menu 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuMainMenu 
      Caption         =   "&Menu"
      Begin VB.Menu GoToParentMenu 
         Caption         =   "&Go to parent."
         Shortcut        =   ^P
      End
      Begin VB.Menu ToggleStatusMenu 
         Caption         =   "&Toggle status."
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu OptionsMainMenu 
      Caption         =   "&Options"
      Begin VB.Menu RetrieveSystemMenusMenu 
         Caption         =   "&Retrieve system menus."
         Shortcut        =   ^S
      End
      Begin VB.Menu RefreshDisplayMenu 
         Caption         =   "&Refresh display."
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "MenuExplorerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface window.
Option Explicit

'This procedure is executed when this window is opened.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   
   With App
      Me.Caption = .Title & ", v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - " & App.CompanyName
   End With
   
   RetrieveSystemMenusMenu.Checked = GetSystemMenus()
   GetWindowList MenuWindowListBox
   
   MenuListBox.Row = 0
   MenuListBox.Col = 0: MenuListBox.Text = "Menu Text:": MenuListBox.ColAlignment(0) = flexAlignLeftCenter
   MenuListBox.Col = 1: MenuListBox.Text = "Menu State:": MenuListBox.ColAlignment(1) = flexAlignLeftCenter
   
   MenuWindowListBox.Row = 0
   MenuWindowListBox.Col = 0: MenuWindowListBox.Text = "Window WindowH:": MenuWindowListBox.ColAlignment(0) = flexAlignRightCenter
   MenuWindowListBox.Col = 1: MenuWindowListBox.Text = "Window Text:": MenuWindowListBox.ColAlignment(1) = flexAlignLeftCenter
   MenuWindowListBox.Col = 2: MenuWindowListBox.Text = "Window Process:": MenuWindowListBox.ColAlignment(2) = flexAlignLeftCenter
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure adjusts this window's controls to its new size.
Private Sub Form_Resize()
On Error Resume Next
   MenuAncestorsBox.Top = Me.ScaleHeight - 1.5
   MenuAncestorsBox.Width = Me.ScaleWidth - 2
   
   MenuListBox.Left = (Me.ScaleWidth / 2) + 1
   MenuListBox.Width = (Me.ScaleWidth / 2) - 2
   MenuListBox.Height = Me.ScaleHeight - 2.5
   MenuListBox.ColWidth(0) = ((MenuListBox.Width * 6) * Screen.TwipsPerPixelX) / MenuListBox.Cols
   MenuListBox.ColWidth(1) = ((MenuListBox.Width * 9) * Screen.TwipsPerPixelX) / MenuListBox.Cols
   
   MenuWindowListBox.Width = (Me.ScaleWidth / 2) - 2
   MenuWindowListBox.Height = Me.ScaleHeight - 2.5
   MenuWindowListBox.ColWidth(0) = ((MenuListBox.Width * 3) * Screen.TwipsPerPixelX) / MenuWindowListBox.Cols
   MenuWindowListBox.ColWidth(1) = ((MenuListBox.Width * 8) * Screen.TwipsPerPixelX) / MenuWindowListBox.Cols
   MenuWindowListBox.ColWidth(2) = ((MenuListBox.Width * 12) * Screen.TwipsPerPixelX) / MenuWindowListBox.Cols
End Sub

'This procedure gives the command to exit a sub menu.
Private Sub GoToParentMenu_Click()
On Error GoTo ErrorTrap
   CurrentMenuH NewCurrentMenuH:=LeaveSubMenu()
   DisplayMenuItems CurrentMenuH(), MenuListBox
   MenuAncestorsBox.Text = GetAncestors()
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub




'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   MsgBox App.Comments, vbInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure enters the selected sub menu if there is one.
Private Sub MenuListBox_DblClick()
On Error GoTo ErrorTrap
   If MenuListBox.Row > 0 Then
      CurrentMenuH NewCurrentMenuH:=EnterSubMenu(CurrentMenuH(), MenuListBox.Row - 1)
      DisplayMenuItems CurrentMenuH(), MenuListBox
      MenuAncestorsBox.Text = GetAncestors()
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to display the menus for the selected window.
Private Sub MenuWindowListBox_RowColChange()
On Error GoTo ErrorTrap

   If MenuWindowListBox.Row > 0 Then
      CurrentMenuH NewCurrentMenuH:=GetWindowMenuH(WindowsH(MenuWindowListBox.Row - 1))
      DisplayMenuItems CurrentMenuH(), MenuListBox
      MenuAncestorsBox.Text = GetAncestors()
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure refreshes the list of menu windows.
Private Sub RefreshDisplayMenu_Click()
On Error GoTo ErrorTrap
   GetWindowList MenuWindowListBox
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to toggle whether or not system menus are retrieved.
Private Sub RetrieveSystemMenusMenu_Click()
On Error GoTo ErrorTrap
   RetrieveSystemMenusMenu.Checked = GetSystemMenus(Toggle:=True)
   GetWindowList MenuWindowListBox
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to disable or enable a menu.
Private Sub ToggleStatusMenu_Click()
On Error GoTo ErrorTrap

   If MenuListBox.Row > 0 Then
      ToggleMenuEnabledState CurrentMenuH(), MenuListBox.Row - 1
      DisplayMenuItems CurrentMenuH(), MenuListBox
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

