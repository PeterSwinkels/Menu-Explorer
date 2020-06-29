Attribute VB_Name = "MenuExplorerModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants used by this program:
Private Const ERROR_SUCCESS As Long = &H0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const INVALID_MENU_HANDLE As Long = &H579&
Private Const MAX_PATH As Long = &H104&
Private Const MAX_STRING As Long = &HFFFF&
Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_DISABLED As Long = &H2&
Private Const MF_ENABLED As Long = &H0&
Private Const MF_GRAYED As Long = &H1&
Private Const MF_SEPARATOR As Long = &H800&
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const WM_GETTEXT As Long = &HD&
Private Const WM_GETTEXTLENGTH As Long = &HE&

'The Microsoft Windows API functions used by this program:
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function EnumWindows Lib "User32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetMenu Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "User32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuState Lib "User32.dll" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuStringA Lib "User32.dll" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetProcessImageFileNameW Lib "Psapi.dll" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetSubMenu Lib "User32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSystemMenu Lib "User32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsMenu Lib "User32.dll" (ByVal hMenu As Long) As Long
Private Declare Function IsWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ModifyMenuA Lib "User32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function SendMessageA Lib "User32.dll" (ByVal hwnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long

'The constants, structures, and variables used by this program:

'This structure contains information about a menu's ancestors.
Type AncestorStr
   AncestorH As Long 'Contains the ancestor menu handles.
   Position As Long  'Contains the ancestor menu positions.
End Type

Private Const NO_API_HANDLE As Long = 0   'Defines a null handle for API functions.
Private Const NO_HANDLE As Long = -1      'Defines a null handle for this program's functions.

Public WindowsH() As Long             'Contains the list of handles of windows that contain menu's.
Private Ancestors() As AncestorStr    'Contains the stack of ancestor menu's for the active sub menu.

'This procedure appends the specified state description to specified text and inserts delimiters.
Private Function AppendStateText(Text As String, State As String) As String
On Error GoTo ErrorTrap
Dim Result As String

   Result = Text
   If Not Result = vbNullString Then Result = Result & " | "
   Result = Result & State
   
EndRoutine:
   AppendStateText = Result
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure checks whether any API errors that have occurred and handles them.
Private Function CheckForError(ReturnValue As Long, Optional Ignored As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue)
      MsgBox Message, vbExclamation
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns/stores the handle for the selected menu.
Public Function CurrentMenuH(Optional NewCurrentMenuH As Long = NO_HANDLE) As Long
On Error GoTo ErrorTrap
Static MenuH As Long

   If Not NewCurrentMenuH = NO_HANDLE Then MenuH = NewCurrentMenuH
   
EndRoutine:
   CurrentMenuH = MenuH
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure displays the items contained by the specified menu in the specified list.
Public Sub DisplayMenuItems(MenuH As Long, Target As MSFlexGrid)
On Error GoTo ErrorTrap
Dim Position As Long
Dim PreviousRow As Long
Dim Text As String

   With Target
      PreviousRow = .Row
      .Rows = 1
         For Position = 0 To CheckForError(GetMenuItemCount(MenuH)) - 1
         .Rows = .Rows + 1
         .Row = .Rows - 1
                    
         Text = Trim$(Replace(GetMenuText(MenuH, Position), vbTab, " "))
         If Text = vbNullString Then Text = "-"
         .Col = 0: .CellFontBold = CBool(IsMenu(CheckForError(GetSubMenu(MenuH, Position)))): .Text = Text
         .Col = 1: .Text = GetMenuStateText(MenuH, Position)
      Next Position
      If PreviousRow < .Rows Then .Row = PreviousRow Else .Row = .Rows - 1
      If .Row > 0 Then .TopRow = .Row
   End With
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the handle for the sub menu at the specified position and updates the ancestor stack.
Public Function EnterSubMenu(ParentMenuH As Long, Position As Long) As Long
On Error GoTo ErrorTrap
Dim SubMenuH As Long

   SubMenuH = NO_API_HANDLE
   If CBool(IsMenu(ParentMenuH)) Then
      SubMenuH = CheckForError(GetSubMenu(ParentMenuH, Position))
      If CBool(IsMenu(SubMenuH)) Then
         Ancestors(UBound(Ancestors())).AncestorH = ParentMenuH
         Ancestors(UBound(Ancestors())).Position = Position
         ReDim Preserve Ancestors(LBound(Ancestors()) To UBound(Ancestors()) + 1) As AncestorStr
      Else
         SubMenuH = ParentMenuH
      End If
   End If
   
EndRoutine:
   EnterSubMenu = SubMenuH
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure returns a list of a menu's ancestors.
Public Function GetAncestors() As String
Dim Index As Long
Dim Result As String

   Result = vbNullString
   For Index = LBound(Ancestors()) To UBound(Ancestors()) - 1
      Result = Result & "\" & GetMenuText(Ancestors(Index).AncestorH, Ancestors(Index).Position)
   Next Index
   
   GetAncestors = Result
End Function

'This procedure returns/sets the flag indicating whether system menus are retrieved.
Public Function GetSystemMenus(Optional Toggle As Boolean = False) As Boolean
On Error GoTo ErrorTrap
Static SystemMenus As Boolean

   If Toggle Then SystemMenus = Not SystemMenus
   
EndRoutine:
   GetSystemMenus = SystemMenus
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure returns the menu handle for the specified window and clears the ancestor list.
Public Function GetWindowMenuH(WindowH As Long) As Long
On Error GoTo ErrorTrap
Dim MenuH As Long

   MenuH = NO_API_HANDLE
   If CBool(IsWindow(WindowH)) Then
      ReDim Ancestors(0 To 0) As AncestorStr
      
      If GetSystemMenus() Then
         MenuH = CheckForError(GetSystemMenu(WindowH, CLng(False)), INVALID_MENU_HANDLE)
      Else
         MenuH = CheckForError(GetMenu(WindowH), INVALID_MENU_HANDLE)
      End If
   End If
   
EndRoutine:
   GetWindowMenuH = MenuH
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the state descriptions for the menu at the specified position.
Private Function GetMenuStateText(MenuH As Long, Position As Long) As String
On Error GoTo ErrorTrap
Dim MenuState As Long
Dim Text As String

   Text = vbNullString
   If CBool(IsMenu(MenuH)) Then
      MenuState = CheckForError(GetMenuState(MenuH, Position, MF_BYPOSITION))
      
      If (MenuState And MF_DISABLED) = MF_DISABLED Then
         Text = AppendStateText(Text, "DISABLED")
      Else
         Text = AppendStateText(Text, "ENABLED")
      End If
      If (MenuState And MF_GRAYED) = MF_GRAYED Then Text = AppendStateText(Text, "GRAYED")
      If (MenuState And MF_SEPARATOR) = MF_SEPARATOR Then Text = AppendStateText(Text, "SEPARATOR")
   End If
   
EndRoutine:
   GetMenuStateText = Text
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the text for the menu at the specified position.
Private Function GetMenuText(MenuH As Long, Position As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim Text As String

   Text = vbNullString
   If CBool(IsMenu(MenuH)) Then
      Text = String$(MAX_STRING, vbNullChar)
      Length = CheckForError(GetMenuStringA(MenuH, Position, Text, Len(Text), MF_BYPOSITION))
      Text = Left$(Text, Length)
   End If
   
EndRoutine:
   GetMenuText = Text
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure retrieves a list of active windows that contain a menu.
Public Sub GetWindowList(Target As MSFlexGrid)
On Error GoTo ErrorTrap
Dim Index As Long
Dim PreviousRow As Long

   ReDim WindowsH(0 To 0) As Long
   
   CheckForError EnumWindows(AddressOf WindowHandler, CLng(0)), INVALID_MENU_HANDLE
   
   With Target
      PreviousRow = .Row
      .Rows = 1
      For Index = LBound(WindowsH()) To UBound(WindowsH()) - 1
         .Rows = .Rows + 1
         .Row = .Rows - 1
         .Col = 0: .Text = CStr(WindowsH(Index))
         .Col = 1: .Text = GetWindowText(WindowsH(Index))
         .Col = 2: .Text = GetWindowProcessImageName(WindowsH(Index))
      Next Index
      If PreviousRow < .Rows Then .Row = PreviousRow Else .Row = .Rows - 1
      If .Row > 0 Then .TopRow = .Row
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure returns the image name for the process owning the specified window.
Private Function GetWindowProcessImageName(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim ImageName As String
Dim Length As Long
Dim ProcessH As Long
Dim ProcessId As Long
 
   ImageName = vbNullString
   CheckForError GetWindowThreadProcessId(WindowH, ProcessId)
   ProcessH = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(False), ProcessId))
   If Not ProcessH = NO_API_HANDLE Then
      ImageName = String$(MAX_PATH, vbNullChar)
      Length = CheckForError(GetProcessImageFileNameW(ProcessH, StrPtr(ImageName), Len(ImageName)))
      CheckForError CloseHandle(ProcessH)
      ImageName = Left$(ImageName, Length)
   End If
   
EndRoutine:
   GetWindowProcessImageName = ImageName
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure returns the text for the specified window.
Private Function GetWindowText(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim Text As String

   Text = String$(CheckForError(SendMessageA(WindowH, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0))) + 1, vbNullChar)
   Length = CheckForError(SendMessageA(WindowH, WM_GETTEXT, ByVal Len(Text), ByVal Text))
   
EndRoutine:
   GetWindowText = Left$(Text, Length)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   
   On Error GoTo ErrorTrap
   MsgBox Err.Description & vbCr & "Error code: " & CStr(ErrorCode), vbExclamation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub
'This procedure returns the handle for the active menu's parent and updates the ancestor stack.
Public Function LeaveSubMenu() As Long
On Error GoTo ErrorTrap
Dim ParentMenuH As Long

   ParentMenuH = CurrentMenuH()
   If Abs(UBound(Ancestors()) - LBound(Ancestors())) > 0 Then
      ParentMenuH = Ancestors(UBound(Ancestors()) - 1).AncestorH
      ReDim Preserve Ancestors(LBound(Ancestors()) To UBound(Ancestors()) - 1) As AncestorStr
   End If
   
EndRoutine:
   LeaveSubMenu = ParentMenuH
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
   CurrentMenuH NewCurrentMenuH:=NO_HANDLE
   
   MenuExplorerWindow.Show
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure disables/enables the menu at the specified position.
Public Sub ToggleMenuEnabledState(MenuH As Long, Position As Long)
On Error GoTo ErrorTrap
Dim MenuState As Long
Dim NewState As Long

   If CBool(IsMenu(MenuH)) Then
      MenuState = CheckForError(GetMenuState(MenuH, Position, MF_BYPOSITION))
      If (MenuState And MF_DISABLED) = MF_DISABLED Then
         NewState = MF_ENABLED
      ElseIf (MenuState And MF_ENABLED) = MF_ENABLED Then
         NewState = MF_DISABLED
      End If
   
      CheckForError ModifyMenuA(MenuH, Position, NewState Or MF_BYPOSITION, CLng(0), GetMenuText(MenuH, Position))
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure checks whether the detected window has a menu.
Private Function WindowHandler(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
Dim MenuH As Long

   If GetSystemMenus() Then
      MenuH = CheckForError(GetSystemMenu(hwnd, CLng(False)), INVALID_MENU_HANDLE)
   Else
      MenuH = CheckForError(GetMenu(hwnd), INVALID_MENU_HANDLE)
   End If
   
   If CBool(IsMenu(MenuH)) Then
      WindowsH(UBound(WindowsH())) = hwnd
      ReDim Preserve WindowsH(LBound(WindowsH()) To UBound(WindowsH()) + 1) As Long
   End If
   
EndRoutine:
   WindowHandler = CLng(True)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

