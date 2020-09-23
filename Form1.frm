VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Generate GUIDs"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "GUID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type GUID
    ab As Long
    ac As Integer
    ad As Integer
    ae(7) As Byte
End Type
      
Private Declare Function CoCreateGuid Lib "OLE32.DLL" _
(ptrGuid As GUID) As Long

Public Function GUID() As String

    Dim lNHI As Long
    Dim udtGuid As GUID
    Dim ab As String
    Dim ac As String
    Dim ad As String
    Dim ae As String
    Dim DLen As Integer
    Dim SLen As Integer
    Dim CLen As Integer
    Dim Tostring As String
    
On Error GoTo errors
    
    Tostring = ""
    
    lNHI = CoCreateGuid(udtGuid)
    
    If lNHI = 0 Then
        
        '1
        ab = Hex$(udtGuid.ab)
        SLen = Len(ab)
        DLen = Len(udtGuid.ab)
        ab = String((DLen * 2) - SLen, "0") & Trim$(ab)
        
        '2
        ac = Hex$(udtGuid.ac)
        SLen = Len(ac)
        DLen = Len(udtGuid.ac)
        ac = String((DLen * 2) - SLen, "0") & Trim$(ac)
        
        '3
        ad = Hex$(udtGuid.ad)
        SLen = Len(ad)
        DLen = Len(udtGuid.ad)
        ad = String((DLen * 2) - SLen, "0") & Trim$(ad)
        
        '4
       For CLen = 0 To 7
       ae = ae & Format$(Hex$(udtGuid.ae(CLen)), "00")
       Next
       Tostring = "{" & ab & "-" & ac & "-" & ad & "-" & ae & "}"
       
       'Tostring = ab & "-" & ac & "-" & ad & "-" & ae & 'With out {}
       'Tostring = ab & ac " & ad  & ae & 'With out -
       
       End If
           
        GUID = Tostring
        
Exit Function

errors:
   MsgBox "Cannot create GUI", vbCritical, "Error"

End Function

Private Sub Form_Load()

   Me.Label2 = GUID()

End Sub
