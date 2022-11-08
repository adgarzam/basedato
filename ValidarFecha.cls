VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidarFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarControlPorValidar As Object 'copia local
Private mvarEsValido As Boolean 'copia local
Private mvarGrupo As String 'copia local
Private mvarMensaje As String 'copia local
Private mvarNombreColoquial As String 'copia local
Private mvarProceder As Boolean 'copia local
Private mvarTipoDeMensaje As TipoMensaje 'copia local
Private mvarValidado As Boolean 'copia local
Private strValorAnterior As String
Private mvarExigirDatos As Boolean 'copia local
Private mvarOmitirBlancosEnExtremos As Boolean 'copia local
Public Property Let OmitirBlancosEnExtremos(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.OmitirBlancosEnExtremos = 5
    mvarOmitirBlancosEnExtremos = vData
End Property


Public Property Get OmitirBlancosEnExtremos() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.OmitirBlancosEnExtremos
    OmitirBlancosEnExtremos = mvarOmitirBlancosEnExtremos
End Property



Public Property Let ExigirDatos(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ExigirDatos = 5
    mvarExigirDatos = vData
End Property


Public Property Get ExigirDatos() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ExigirDatos
    ExigirDatos = mvarExigirDatos
End Property




Public Function Validar() As Boolean
  strValorAnterior = Me.ControlPorValidar.Text
  
  If Me.OmitirBlancosEnExtremos Then
    strValorAnterior = Trim(strValorAnterior)
  End If
  
  If Len(strValorAnterior) > 0 Then
    Me.EsValido = IsDate(strValorAnterior)
  Else
    Me.EsValido = Not Me.ExigirDatos
  End If
  
  Me.Validado = True
  Validar = Me.EsValido
End Function

Public Property Let Validado(ByVal vData As Boolean)
    mvarValidado = vData
End Property

Public Property Get Validado() As Boolean
    If Me.OmitirBlancosEnExtremos Then
      Me.Validado = (strValorAnterior = Trim(Me.ControlPorValidar.Text))
    Else
      Me.Validado = (strValorAnterior = Me.ControlPorValidar.Text)
    End If
    Validado = mvarValidado
End Property

Public Property Let TipoDeMensaje(ByVal vData As TipoMensaje)
    mvarTipoDeMensaje = vData
End Property

Public Property Get TipoDeMensaje() As TipoMensaje
  TipoDeMensaje = mvarTipoDeMensaje
End Property

Public Property Let Proceder(ByVal vData As Boolean)
    mvarProceder = vData
End Property

Public Property Get Proceder() As Boolean
    Proceder = mvarProceder
End Property

Public Property Let NombreColoquial(ByVal vData As String)
    mvarNombreColoquial = vData
End Property

Public Property Get NombreColoquial() As String
    NombreColoquial = mvarNombreColoquial
End Property

Public Property Let Mensaje(ByVal vData As String)
    If vData <> "" Then
      mvarMensaje = vData
    End If
End Property

Public Property Get Mensaje() As String
    Mensaje = mvarMensaje
End Property

Public Property Let Grupo(ByVal vData As String)
    If vData <> "" Then
      mvarGrupo = vData
    End If
End Property

Public Property Get Grupo() As String
    Grupo = mvarGrupo
End Property

Public Property Let EsValido(ByVal vData As Boolean)
    mvarEsValido = vData
End Property

Public Property Get EsValido() As Boolean
   If Not Me.Validado Then
     Me.Validar
   End If
   EsValido = mvarEsValido
End Property

Public Property Set ControlPorValidar(ByVal vData As Object)
    ' El control por validar debe contar con una
    '  propiedad Text.
    ' Var
    Dim strTipo As String
    
   If mvarControlPorValidar Is Nothing Then
      On Error Resume Next
      strTipo = vData.Text
      If Err.Number = 0 Then
        On Error GoTo 0
        Set mvarControlPorValidar = vData
      Else
        MsgBox "Este validador sólo admite controles que cuenten con una propiedad Text.", vbCritical, "Error " & Err.Number
        Err.Clear
        On Error GoTo 0
      End If
   Else
     MsgBox "Sólo se puede asignar una vez el control.", vbExclamation, "Error"
   End If
End Property

Public Property Get ControlPorValidar() As Object
    Set ControlPorValidar = mvarControlPorValidar
End Property

Private Sub Class_Initialize()
  mvarMensaje = "El campo no tiene un formato correcto de fecha."
  mvarTipoDeMensaje = 48
  mvarGrupo = "General"
  mvarProceder = True
  mvarExigirDatos = True
  mvarOmitirBlancosEnExtremos = True
End Sub
