VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidarExpresiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ValidarExpresion"
Attribute VB_Ext_KEY = "Member0" ,"ValidarExpresion"
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

Public Function Add(ControlPorValidar As Object, Optional TipoExpresion As TiposExpresion = TiposExpresion.General, Optional Expresion As String = "*", Optional Mensaje As String = "", Optional NombreColoquial As String = "", Optional Grupo As String = "General", Optional Key As String = "") As ValidarExpresion
    'crear un nuevo objeto
    Dim objNewMember As ValidarExpresion
    Set objNewMember = New ValidarExpresion


    ' Establecer las propiedades que se transfieren al método
    Set objNewMember.ControlPorValidar = ControlPorValidar
    
    If TipoExpresion <> TiposExpresion.General Then
      objNewMember.TipoExpresion = TipoExpresion
    ElseIf Expresion <> "*" Then
      objNewMember.Expresion = Expresion
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

Public Property Get Item(vntIndexKey As Variant) As ValidarExpresion
Attribute Item.VB_UserMemId = 0
    'se usa al hacer referencia a un elemento de la colección
    'vntIndexKey contiene el índice o la clave de la colección,
    'por lo que se declara como un Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
  Err.Clear
  On Error GoTo 0
End Property



Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la
    'colección. Sintaxis: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'se usa al quitar un elemento de la colección
    'vntIndexKey contiene el índice o la clave, por lo que se
    'declara como un Variant
    'Sintaxis: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'crea la colección cuando se crea la clase
    Set mCol = New Collection
    mvarProceder = True
End Sub


Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
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

