Attribute VB_Name = "Mod_ShowDirs"
'.=========================================================================
'.Browse Folders Module
'.Copyright 1999 Tribble Software.  All rights reserved.
'.Phone        : (616) 455-2055
'.E-mail       : carltribble@earthlink.net
'.=========================================================================
' DO NOT DELETE THE COMMENTS ABOVE.  All other comments in this module
' may be deleted from production code, but lines above must remain.
'--------------------------------------------------------------------------
'.Description  : This module calls three functions in shell32.dll to allow
'.               the user to browse for a folder.
'.
'.Written By   : Carl Tribble
'.Date Created : 10/05/1999 08:06:31 PM
'.Rev. History :
' Comments     : The public entry point is the procedure tsGetPathFromUser,
'                The selected folder name is returned in the form of a full
'                path but without the trailing "\". If the User presses
'                Cancel, or an error occurs, the procedure returns Null.
'                This module is completely self-contained.  Simply copy it
'                into your database to use it.
'.-------------------------------------------------------------------------
'.
' ADDITIONAL NOTES:
'
'  If you want your user to browse for file names you must use the module
'  basBrowseFiles instead, or the common dialog activeX control.
'
'  TO STREAMLINE this module for production programs, you should remove:
'     1) Unnecessary comments
'     2) Flag and Root Folder Constants which you do not intend to use.
'     3) The test procedure tsGetPathFromUserTest
'       *DO NOT REMOVE ANYTHING ELSE. Everything else is required.
'
'--------------------------------------------------------------------------
'
' INSTRUCTIONS:
'
'         ( For a working example, open the Debug window  )
'         ( and enter tsGetPathFromUserTest.              )
'         (                                               )
'         ( frmBrowseFoldersTest, if available, provides  )
'         ( additional testing features.                  )
'
'.All the arguments for the function are optional.  You may call it with no
'.arguments whatsoever and simply assign its return value to a variable of
'.the Variant type.  For example:
'.
'.   varFileName = tsGetPathFromUser()
'.
'.The function will return:
'.   the full path selected by the user, or
'.   Null if an error occurs or if the user presses Cancel.
'.
'.Optional arguments may include any of the following:
'. rlngFlags     : one or more of the tscBF* Flag constants (declared
'.                 below). Combine multiple constants like this:
'.                   tscBFReturnOnlyFSDirs Or tscBFDontGoBelowDomain
'. lngRootFolder : a tscRF Root Folder constant (declared below) indicating
'.                 what folder you want to start with.  These constants are
'.                 not to be combined, just pick the one you want to use.
'. strHeaderMsg  : a message you want to appear at the top of the dialog
'.                 box.  Note although it is refered to internally as the
'.                 Title it is NOT the dialog title, aka caption (the
'                  caption is always "Browse for Folder").  The message
'                  can be up to about 110 characters in length and
'                  up to two lines.  It appears below the Title bar, but
'                  above the actual folder box.
'
'.-------------------------------------------------------------------------
'.
Option Explicit

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
 Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
 ByVal pszPath As String) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
 (ByVal hwndOwner As Long, ByVal nFolder As Long, _
 pidl As ITEMIDLIST) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
 "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Private Type SHITEMID
   cb As Long
   abID As Byte
End Type

Private Type ITEMIDLIST
   mkid As SHITEMID
End Type

' Flag Constants
Public Const tscBFReturnOnlyFSDirs = &H1
Public Const tscBFDontGoBelowDomain = &H2
Public Const tscBFStatusText = &H4
Public Const tscBFReturnFSAncestors = &H8
Public Const tscBFBrowseForComputer = &H1000
Public Const tscBFBrowseForPrinter = &H2000

' Root Folder Constants
Public Const tscRFDesktop = &H0
Public Const tscRFPrograms = &H2
Public Const tscRFControls = &H3
Public Const tscRFPrinters = &H4
Public Const tscRFPersonal = &H5
Public Const tscRFFavorites = &H6
Public Const tscRFRecent = &H8
Public Const tscRFBitBucket = &HA
Public Const tscRFDesktopDirectory = &H10
Public Const tscRFDrives = &H11
Public Const tscRFNetwork = &H12
Public Const tscRFNethood = &H13
Public Const tscRFTemplates = &H15

Public Function tsGetPathFromUser( _
 Optional ByRef rlngflags As Long = tscBFReturnOnlyFSDirs, _
 Optional ByVal lngRootFolder As Long = tscRFDrives, _
 Optional ByVal strHeaderMsg As String = "") As Variant
   
   On Error GoTo tsGetPathFromUser_Err
   Const conBufLen = 512
   Dim bi As BROWSEINFO
   Dim idl As ITEMIDLIST
   Dim lngReturn As Long
   Dim pidl As Long
   Dim strPath As String

   bi.hOwner = 0
   lngReturn = SHGetSpecialFolderLocation( _
    ByVal bi.hOwner, lngRootFolder, idl)
   bi.pidlRoot = idl.mkid.cb
   bi.lpszTitle = strHeaderMsg
   bi.ulFlags = rlngflags
   pidl = SHBrowseForFolder(bi)
   strPath = Space(conBufLen)
   lngReturn = SHGetPathFromIDList(ByVal pidl, ByVal strPath)
   
   If lngReturn <> 0 Then
      tsGetPathFromUser = tsTrimNull(strPath)
   Else
      tsGetPathFromUser = ""
   End If
   
tsGetPathFromUser_End:
   On Error GoTo 0
   Exit Function

tsGetPathFromUser_Err:
   Beep
   MsgBox Err.Description, , "Error: " & Err.Number _
    & " in function basBrowseFolders.tsGetPathFromUser"
   Resume tsGetPathFromUser_End

End Function

' Trim Nulls from a string returned by an API call.

Private Function tsTrimNull(ByVal strItem As String) As String
   
   On Error GoTo tsTrimNull_Err
   Dim i As Integer
   
   i = InStr(strItem, vbNullChar)
   If i > 0 Then
       tsTrimNull = Left(strItem, i - 1)
   Else
       tsTrimNull = strItem
   End If
    
tsTrimNull_End:
   On Error GoTo 0
   Exit Function

tsTrimNull_Err:
   Beep
   MsgBox Err.Description, , "Error: " & Err.Number _
    & " in function basBrowseFolders.tsTrimNull"
   Resume tsTrimNull_End

End Function

'--------------------------------------------------------------------------
' Project      : tsDeveloperTools
' Description  : An example of how you can call tsGetPathFromUser()
' Calls        :
' Accepts      :
' Returns      :
' Written By   : Carl Tribble
' Date Created : 05/04/1999 11:19:41 AM
' Rev. History :
' Comments     : This is provided merely as an example to the programmer
'                It may be safely deleted from production code.
'--------------------------------------------------------------------------

Public Sub tsGetPathFromUserTest()
   
   On Error GoTo tsGetPathFromUserTest_Err
   Dim lngFlags As Long
   Dim lngRoot As Long
   Dim strHeaderMsg As String
   Dim varPath As Variant
   
   lngFlags = tscBFReturnOnlyFSDirs Or tscBFDontGoBelowDomain
   lngRoot = tscRFDrives
   strHeaderMsg = "This is where the header message displays. " _
    & vbCrLf & "Note it only holds 2 full lines (about 100 " _
    & "characters altogether)."
   varPath = tsGetPathFromUser(lngFlags, lngRoot, strHeaderMsg)

   If IsNull(varPath) Then
      Debug.Print "User pressed 'Cancel'."
   Else
      Debug.Print varPath
   End If

tsGetPathFromUserTest_End:
   On Error GoTo 0
   Exit Sub

tsGetPathFromUserTest_Err:
   Beep
   MsgBox Err.Description, , "Error: " & Err.Number _
    & " in sub basBrowseFolders.tsGetPathFromUserTest"
   Resume tsGetPathFromUserTest_End

End Sub


