VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 20 januari 2009
' 3:56 AM
'=======================================
' cFileInfo
'=======================================
Option Explicit

Dim varFileName As String
Dim varFilePath As String
Dim varFileSize As Double
Dim varFileTitle As String
Dim varFileAttr As Integer

Dim varVirusID As String
Dim varVirusID_DET As String
Dim varVirusName As String
Dim varVirusType As String
Dim varVirusAlias As String
Dim varVirusWilds As String
Dim varVirusDamage As String
Dim varVirusDistrb As String
Dim varVirusCRC As String
Dim varVirusClean As String
Dim varVirusAction As String

Property Let VirusAction(value As String)
    varVirusAction = value
End Property

Property Get VirusAction() As String
    VirusAction = varVirusAction
End Property

Property Let VirusClean(value As String)
    varVirusClean = value
End Property

Property Get VirusClean() As String
    VirusClean = varVirusClean
End Property

Property Let VirusID(value As String)
    varVirusID = value
End Property

Property Get VirusID() As String
    VirusID = varVirusID
End Property

Property Let VirusIDDet(value As String)
    varVirusID_DET = value
End Property

Property Get VirusIDDet() As String
    VirusIDDet = varVirusID_DET
End Property

Property Let VirusName(value As String)
    varVirusName = value
End Property

Property Get VirusName() As String
    VirusName = varVirusName
End Property

Property Let VirusType(value As String)
    varVirusType = value
End Property

Property Get VirusType() As String
    VirusType = varVirusType
End Property

Property Let VirusAlias(value As String)
    varVirusAlias = value
End Property

Property Get VirusAlias() As String
    VirusAlias = varVirusAlias
End Property

Property Let VirusWilds(value As String)
    varVirusWilds = value
End Property

Property Get VirusWilds() As String
    VirusWilds = varVirusWilds
End Property

Property Let VirusDamage(value As String)
    varVirusDamage = value
End Property

Property Get VirusDamage() As String
    VirusDamage = varVirusDamage
End Property

Property Let VirusDistrb(value As String)
    varVirusDistrb = value
End Property

Property Get VirusDistrb() As String
    VirusDistrb = varVirusDistrb
End Property

Property Let VirusCRC(value As String)
    varVirusCRC = value
End Property

Property Get VirusCRC() As String
    VirusCRC = varVirusCRC
End Property

Property Let FileName(value As String)
    varFileName = value
End Property

Property Get FileName() As String
    FileName = varFileName
End Property

Property Let FilePath(value As String)
    varFilePath = value
End Property

Property Get FilePath() As String
    FilePath = varFilePath
End Property

Property Let FileSize(value As Double)
    varFileSize = value
End Property

Property Get FileSize() As Double
    FileSize = varFileSize
End Property

Property Let FileTitle(value As String)
    varFileTitle = value
End Property

Property Get FileTitle() As String
    FileTitle = varFileTitle
End Property

Property Let FileAttr(value As VbFileAttribute)
    varFileAttr = value
End Property

Property Get FileAttr() As VbFileAttribute
    FileAttr = varFileAttr
End Property

