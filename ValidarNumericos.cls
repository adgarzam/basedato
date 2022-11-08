VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidarNumericos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ValidarNumerico"
Attribute VB_Ext_KEY = "Member0" ,"ValidarNumerico"
Option Explicit

'variable local para contener colección
Private mCol As Collection
Private mvarProceder As Boolean 'copia local
Public Property Let Proceder(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Proceder = 5
    mvarProceder = vData
End Property


Public Property Get Proceder() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Proceder
    Proceder = mvarProceder
End Property




Public Sub Clear()
  While mCol.Count > 0
    mCol.Remove 1
  Wend
End Sub

Public Function Add(ControlPorValidar As Object, Optional Comparacion As TipoNumero = Indistinto, Optional Mensaje As String = "", Optional NombreColoquial As String, Optional Grupo As String = "General", Optional Key As String = "") As ValidarNumerico
    'crear un nuevo objeto
    Dim objNewMember As ValidarNumerico
    Set objNewMember = New ValidarNumerico

    ' Establecer las propiedades que se transfieren al método
    Set objNewMember.ControlPorValidar = ControlPorValidar
    
    If Comparacion <> Indistinto Then
      objNewMember.Comparacion = Comparacion
    End If
    
    If Mensaje <> "" Then
      objNewMember.Mensaje = Mensaje
    End If
    
    If NombreColoquial <> "" Then
      objNewMember.NombreColoquial = NombreColoquial
    Else
      objNewMember.NombreColoquial = ControlPorValidar.Name
    End If
    
    If Grupo <> "General" Then
      objNewMember.Grupo = Grupo
    End If
    
    objNewMember.Key = Key
    
    If Not objNewMember.ControlPorValidar Is Nothing Then
      If Len(Key) = 0 Then
          mCol.Add objNewMember
      Else
          mCol.Add objNewMember, Key
      End If
    
      'devolver el objeto creado
      Set Add = objNewMember
    Else
      Set Add = Nothing
    End If
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ValidarNumerico
Attribute Item.VB_UserMemId = 0
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
  Err.Clear
  On Error GoTo 0
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    mvarProceder = True
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Function Existe(ByVal CualElemento As Variant) As Boolean
  ' Var
  Dim Elem As Object
  
    On Error Resume Next
  Set Elem = mCol(CualElemento)
  Err.Clear
  On Error GoTo 0
  Existe = Not (Elem Is Nothing)
End Function

