Attribute VB_Name = "Mod_Declarations"
Option Explicit

'Global declaration
Public PackFileType As Byte                         'Type of file (Zip, GZ.etc..etc)
Public PackTotFiles As Long
Public PackFileName As String
Public PackComments As String
Public PackData() As Byte                       'Used to store Packed/Unpacked data

'signatures short version
Public Const ZipHeader As Integer = &H4B50      'PK-zip signature
Public Const GZipHeader As Integer = &H8B1F     'Gzip signature
Public Const ZHeader As Integer = &H9D1F        'Z signature
Public Const PackHeader As Integer = &H1E1F     'Pack signature
Public Const FreezeHeader As Integer = &HA21F   'Freeze signature
Public Const CABHeader As Integer = &H534D      'Cab Signature
Public Const ARJHeader As Integer = &HEA60      'ARJ Signature
Public Const RARHeader As Integer = &H6152      'RAR Signature
Public Const ARCHeader As Integer = &H1A        'ARC header has only one byte

'FileType Declarations
Public Const ZipFileType As Integer = 1
Public Const GZFileType As Integer = 2
Public Const TARFileType As Integer = 3
Public Const ARJFileType As Integer = 4
Public Const LZHFileType As Integer = 5
Public Const RARFileType As Integer = 6
Public Const ARCFileType As Integer = 7
Public Const CABFileType As Integer = 7


