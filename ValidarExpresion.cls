VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidarExpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Mensaje"
Option Explicit

' Enumeraci�n
Public Enum TiposExpresion
    General = 0
    CorreoE = 1
    CodPostal = 2
    Telefono10 = 3
    Telefono9 = 4
    Telefono8 = 5
    Telefono7 = 6
    Telefono6 = 7
    Telefono5 = 8
    TelefonoLD = 9
    RFC = 10
    CURP = 11
    RFCHomo = 12
    Usuario = 13
End Enum

Public Key As String

Private mvarControlPorValidar As Object 'copia local
Private mvarMensaje As String 'copia local
Private mvarExpresion As String 'copia local
Private mvarEsValido As Boolean 'copia local
Private mvarNombreColoquial As String 'copia local
Private mvarComparacion As TipoEvaluacion 'copia local
Private mvarValidado As Boolean 'copia local
Private mvarTipoDeMensaje As TipoMensaje 'copia local
Private mvarGrupo As String 'copia local
Private strValorAnterior As String
Private mvarTipoExpresion As TiposExpresion 'copia local
Private mvarProceder As Boolean 'copia local
Private mvarExigirDatos As Boolean 'copia local
Private mvarOmitirBlancosEnExtremos As Boolean 'copia local
Public Property Let OmitirBlancosEnExtremos(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.OmitirBlancosEnExtremos = 5
    mvarOmitirBlancosEnExtremos = vData
End Property


Public Property Get OmitirBlancosEnExtremos() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.OmitirBlancosEnExtremos
    OmitirBlancosEnExtremos = mvarOmitirBlancosEnExtremos
End Property



Public Property Let ExigirDatos(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.ExigirDatos = 5
    mvarExigirDatos = vData
End Property


Public Property Get ExigirDatos() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.ExigirDatos
    ExigirDatos = mvarExigirDatos
End Property



Public Property Let Proceder(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.Proceder = 5
    mvarProceder = vData
End Property


Public Property Get Proceder() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Proceder
    Proceder = mvarProceder
End Property

Public Property Let TipoExpresion(ByVal vData As TiposExpresion)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci�n Set.
'Syntax: Set x.TipoExpresion = Form1
    Select Case vData
      Case General
        Me.Expresion = "*"
      Case CorreoE
        Me.Expresion = "*@*.???|*@*.???.*"
      Case CodPostal
        Me.Expresion = "#####"
      Case Telefono10
        Me.Expresion = "##-##-##-##-##|## ## ## ## ##|##########"
      Case Telefono9
        Me.Expresion = "#-##-##-##-##|# ## ## ## ##|#########"
      Case Telefono8
        Me.Expresion = "##-##-##-##|## ## ## ##|########"
      Case Telefono7
        Me.Expresion = "#-##-##-##|# ## ## ##|#######"
      Case Telefono6
        Me.Expresion = "##-##-##|## ## ##|######"
      Case Telefono5
        Me.Expresion = "#-##-##|# ## ##|#####"
      Case TelefonoLD
        Me.Expresion = "(##) (##) ##-##-##-##|(##)(##)##-##-##-##|##-##-##-##-##-##|(##) ## ##-##-##-##|############"
      Case RFC
        Me.Expresion = "[A-Z][A-Z][A-Z][A-Z]-######-???|[A-Z][A-Z][A-Z][A-Z] ###### ???|[A-Z][A-Z][A-Z][A-Z]######???|[A-Z][A-Z][A-Z][A-Z]-######|[A-Z][A-Z][A-Z][A-Z] ######|[A-Z][A-Z][A-Z][A-Z]######"
      Case CURP
        Me.Expresion = "[A-Z][A-Z][A-Z][A-Z]######[HM][A-Z][A-Z]?????"
      Case RFCHomo
        Me.Expresion = "[A-Z][A-Z][A-Z][A-Z]-######-???|[A-Z][A-Z][A-Z][A-Z] ###### ???|[A-Z][A-Z][A-Z][A-Z]######???"
      Case Usuario
        Me.Expresion = ""
    End Select
    mvarTipoExpresion = vData
End Property

Public Property Get TipoExpresion() As TiposExpresion
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.TipoExpresion
    TipoExpresion = mvarTipoExpresion
End Property

Public Property Let Grupo(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.Grupo = 5
    mvarGrupo = vData
End Property

Public Property Get Grupo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Grupo
    Grupo = mvarGrupo
End Property

Public Property Let TipoDeMensaje(ByVal vData As TipoMensaje)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.TipoDeMensaje = 5
    mvarTipoDeMensaje = vData
End Property

Public Property Get TipoDeMensaje() As TipoMensaje
  'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
  'Syntax: Debug.Print X.TipoDeMensaje
    TipoDeMensaje = mvarTipoDeMensaje
End Property

Public Function Validar() As Boolean
  ' Var
  Dim a As Variant
  Dim i As Integer
  Dim l As Integer
  Dim u As Integer
  Dim bolCorrecto As Boolean
  
  strValorAnterior = Me.ControlPorValidar.Text
  If Me.OmitirBlancosEnExtremos Then
    strValorAnterior = Trim(strValorAnterior)
  End If
  
  If Len(strValorAnterior) > 0 Then
    a = Split(mvarExpresion, "|")
    l = LBound(a)
    u = UBound(a)
    For i = l To u
      bolCorrecto = strValorAnterior Like a(i)
      If bolCorrecto Then Exit For
    Next i
  Else
    bolCorrecto = Not Me.ExigirDatos
  End If
  
  Me.EsValido = bolCorrecto
  Me.Validado = True
  Validar = Me.EsValido
End Function

Public Property Let Validado(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.Validado = 5
    mvarValidado = vData
End Property

Public Property Get Validado() As Boolean
    If strValorAnterior = Me.ControlPorValidar.Text And Not Me.OmitirBlancosEnExtremos Then
      Me.Validado = True
    ElseIf strValorAnterior = Trim(Me.ControlPorValidar.Text) And Me.OmitirBlancosEnExtremos Then
      Me.Validado = True
    Else
      Me.Validado = False
    End If
    Validado = mvarValidado
End Property

Public Property Let Comparacion(ByVal vData As TipoEvaluacion)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci�n Set.
'Syntax: Set x.Comparacion = Form1
    mvarComparacion = vData
End Property

Public Property Get Comparacion() As TipoEvaluacion
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Comparacion
   Comparacion = mvarComparacion
End Property

Public Property Let NombreColoquial(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.NombreColoquial = 5
    mvarNombreColoquial = vData
End Property

Public Property Get NombreColoquial() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.NombreColoquial
    NombreColoquial = mvarNombreColoquial
End Property

Public Property Let EsValido(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.EsValido = 5
    mvarEsValido = vData
End Property


Public Property Get EsValido() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.EsValido
    EsValido = mvarEsValido
End Property

Public Property Let Expresion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.Expresion = 5
    mvarExpresion = vData
    mvarTipoExpresion = General
End Property

Public Property Get Expresion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Expresion
    Expresion = mvarExpresion
End Property

Public Property Let Mensaje(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci�n.
'Syntax: X.Mensaje = 5
    mvarMensaje = vData
End Property


Public Property Get Mensaje() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Mensaje
    Mensaje = mvarMensaje
End Property

Public Property Set ControlPorValidar(ByVal vData As Object)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci�n Set.
'Syntax: Set x.ControlPorValidar = Form1
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
        MsgBox "Este validador s�lo admite controles que cuenten con una propiedad Text.", vbCritical, "Error " & Err.Number
        Err.Clear
        On Error GoTo 0
      End If
   Else
     MsgBox "S�lo se puede asignar una vez el control.", vbExclamation, "Error"
   End If
End Property

Public Property Get ControlPorValidar() As Object
  Set ControlPorValidar = mvarControlPorValidar
End Property

Private Sub Class_Initialize()
  mvarMensaje = "El valor no cumple con el patr�n dado."
  mvarTipoDeMensaje = 48
  mvarGrupo = "General"
  mvarExpresion = "*"
  mvarProceder = True
  mvarExigirDatos = True
  mvarOmitirBlancosEnExtremos = True
End Sub
