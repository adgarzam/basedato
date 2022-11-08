VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarcaObligatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"MarcaObligatorio"
Attribute VB_Ext_KEY = "Member0" ,"MarcaObligatorio"
Option Explicit

'variable local para contener colección
Private mCol As Collection
Private lngMsj As Long

Public Sub Clear()
  While mCol.Count > 0
    mCol.Remove 1
  Wend
  lngMsj = 0
End Sub

Public Sub Mostrar()
  ' Var
  Dim mob As MarcaObligatorio
  
  For Each mob In mCol
    mob.Mostrar
  Next mob
  
  If mCol.Count > 0 Then
    Beep
  End If
End Sub

Public Function Add(CualControl As Object, Optional Mensaje As String = "", Optional TipoDeMarcado As TipoMarcado = Obligatorio, Optional Key As String) As MarcaObligatorio
    'crear un nuevo objeto
    Dim objNewMember As MarcaObligatorio
    Set objNewMember = New MarcaObligatorio
    
    lngMsj = lngMsj + 1
    
    'establecer las propiedades que se transfieren al método
    If Mensaje = "" Then
'      Select Case TipoDeMarcado   ' Tipo de marcado SIEMPRE estará establecido
'        Case Obligatorio ', TipoMarcado.General
'          Mensaje = "Este campo es obligatorio."
'        Case Optativo
'          Mensaje = "Este campo es optativo."
'        Case Formato
'          Mensaje = "Este campo debe tener datos con el formato adecuado."
'      End Select
    End If
    
    If Key = "" Then
      Key = "Obligatorio" & lngMsj
    End If
    
    objNewMember.Key = Key
    Set objNewMember.CualControl = CualControl
    If Mensaje <> "" Then
      objNewMember.Mensaje = Mensaje
      objNewMember.MensajeCompuesto = Mensaje
    End If
    objNewMember.TipoDeMarcado = TipoDeMarcado
    mCol.Add objNewMember, Key

    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As MarcaObligatorio
Attribute Item.VB_UserMemId = 0
    'se usa al hacer referencia a un elemento de la colección
    'vntIndexKey contiene el índice o la clave de la colección,
    'por lo que se declara como un Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
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
End Sub


Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
    Set mCol = Nothing
End Sub

