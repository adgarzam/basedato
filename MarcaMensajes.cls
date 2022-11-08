VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarcaMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"MarcaMensaje"
Attribute VB_Ext_KEY = "Member0" ,"MarcaMensaje"
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
  Dim mme As MarcaMensaje
  
  For Each mme In mCol
    mme.Mostrar
  Next mme
  
  If mCol.Count > 0 Then
    Beep
  End If
End Sub

Public Function Add(CualControl As Object, Optional Mensaje As String = "", Optional TipoDeMensaje As TipoMensaje = Exclamacion, Optional Key As String = "") As MarcaMensaje
    'crear un nuevo objeto
    Dim objNewMember As MarcaMensaje
    Set objNewMember = New MarcaMensaje

    lngMsj = lngMsj + 1
    
    ' Establece las propiedades optativas
    If Mensaje = "" Then
      Select Case TipoDeMensaje  ' Tipo de mensaje SIEMPRE estará establecido
        Case Informacion
          Mensaje = "Hay observaciones en este control."
        Case Critico
          Mensaje = "Hay un error serio en este control."
        Case Exclamacion
          Mensaje = "Hay una situación de cuidado en este control."
        Case Pregunta
          Mensaje = "Hay dudas respecto a este control."
      End Select
    End If
    
    If Key = "" Then
      Key = "Mensaje" & lngMsj
    End If
    
    'establecer las propiedades que se transfieren al método
    objNewMember.Key = Key
    Set objNewMember.CualControl = CualControl
    objNewMember.Mensaje = Mensaje
    objNewMember.MensajeCompuesto = Mensaje
    objNewMember.TipoDeMensaje = TipoDeMensaje
    mCol.Add objNewMember, Key

    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As MarcaMensaje
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

