VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidarRango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Mensaje"
Option Explicit

Public Key As String
Private mvarControlPorValidar As Object 'copia local
Private mvarMensaje As String 'copia local
Private mvarValorInicial As Variant 'copia local
Private mvarValorFinal As Variant 'copia local
Private mvarEsValido As Boolean 'copia local
Private mvarNombreColoquial As String 'copia local
Private mvarValidado As Boolean 'copia local
Private strValorAnterior As String
Private mvarTipoDeMensaje As TipoMensaje 'copia local
Private mvarGrupo As String 'copia local
Private mvarProceder As Boolean 'copia local
Private mvarTipoDeValores As TipoValores 'copia local
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




Public Property Let TipoDeValores(ByVal vData As TipoValores)
  mvarTipoDeValores = vData
End Property

Public Property Get TipoDeValores() As TipoValores
  TipoDeValores = mvarTipoDeValores
End Property

Public Property Let Proceder(ByVal vData As Boolean)
    mvarProceder = vData
End Property

Public Property Get Proceder() As Boolean
    Proceder = mvarProceder
End Property

Public Property Let Grupo(ByVal vData As String)
    mvarGrupo = vData
End Property

Public Property Get Grupo() As String
    Grupo = mvarGrupo
End Property

Public Property Let TipoDeMensaje(ByVal vData As TipoMensaje)
    mvarTipoDeMensaje = vData
End Property

Public Property Get TipoDeMensaje() As TipoMensaje
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.TipoDeMensaje
    TipoDeMensaje = mvarTipoDeMensaje
End Property

Public Function Validar() As Boolean
  ' Var
  Dim varValor As Variant
  Dim bolContinuar As Boolean
  
  strValorAnterior = Me.ControlPorValidar.Text
  If Me.OmitirBlancosEnExtremos Then
    strValorAnterior = Trim(strValorAnterior)
  End If
  
  If Len(strValorAnterior) > 0 Then
    Select Case Me.TipoDeValores
      Case TipoValores.General
        varValor = Me.ControlPorValidar.Text
        bolContinuar = True
      Case TipoValores.Cadena
        varValor = CStr(Me.ControlPorValidar.Text)
        bolContinuar = True
      Case TipoValores.CadenaIgnorarMayYMin
        varValor = UCase(Me.ControlPorValidar.Text)
        bolContinuar = True
      Case TipoValores.Fechas
        If IsDate(Me.ControlPorValidar.Text) Then
          varValor = CDate(Me.ControlPorValidar.Text)
          bolContinuar = True
        Else
          bolContinuar = False
        End If
      Case TipoValores.Numeros
        If IsNumeric(Me.ControlPorValidar.Text) Then
          varValor = Val(Me.ControlPorValidar.Text)
          bolContinuar = True
        Else
          bolContinuar = False
        End If
    End Select
    
    If bolContinuar Then
      If (Me.ValorInicial <> Me.ValorFinal) Then
        If Me.ValorInicial < Me.ValorFinal Then
          If Me.TipoDeValores = CadenaIgnorarMayYMin Then
            Me.EsValido = ((varValor >= UCase(Me.ValorInicial)) And (varValor <= UCase(Me.ValorFinal)))
          Else
            Me.EsValido = ((varValor >= Me.ValorInicial) And (varValor <= Me.ValorFinal))
          End If
        Else
          If Me.TipoDeValores = CadenaIgnorarMayYMin Then
            Me.EsValido = ((varValor >= UCase(Me.ValorFinal)) And (varValor <= UCase(Me.ValorInicial)))
          Else
            Me.EsValido = ((varValor >= Me.ValorFinal) And (varValor <= Me.ValorInicial))
          End If
        End If
      End If
    Else
      Me.EsValido = False
    End If
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
   If strValorAnterior = Me.ControlPorValidar.Text And Not Me.OmitirBlancosEnExtremos Then
      Me.Validado = True
    ElseIf strValorAnterior = Trim(Me.ControlPorValidar.Text) And Me.OmitirBlancosEnExtremos Then
      Me.Validado = True
    Else
      Me.Validado = False
    End If
    Validado = mvarValidado
End Property

Public Property Let NombreColoquial(ByVal vData As String)
    mvarNombreColoquial = vData
End Property

Public Property Get NombreColoquial() As String
    NombreColoquial = mvarNombreColoquial
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

Public Property Let ValorFinal(ByVal vData As Variant)
    mvarValorFinal = vData
End Property

Public Property Get ValorFinal() As Variant
    ValorFinal = mvarValorFinal
End Property

Public Property Let ValorInicial(ByVal vData As Variant)
    mvarValorInicial = vData
End Property

Public Property Get ValorInicial() As Variant
    ValorInicial = mvarValorInicial
End Property

Public Property Let Mensaje(ByVal vData As String)
    mvarMensaje = vData
End Property

Public Property Get Mensaje() As String
    Mensaje = mvarMensaje
End Property

Public Property Set ControlPorValidar(ByVal vData As Object)
    ' Var
    Dim strTipo As String
    
    If mvarControlPorValidar Is Nothing Then
      strTipo = UCase(TypeName(vData))
      If strTipo = "TEXTBOX" Or strTipo = "COMBOBOX" Then
        Set mvarControlPorValidar = vData
      Else
        MsgBox "Este validador sólo admite controles de tipo TextBox o ComboBox.", vbCritical, "Error"
      End If
    Else
      MsgBox "Sólo se puede asignar una vez el control por validar."
    End If
End Property


Public Property Get ControlPorValidar() As Object
    Set ControlPorValidar = mvarControlPorValidar
End Property

Private Sub Class_Initialize()
  mvarMensaje = "El valor está fuera del rango."
  mvarValorInicial = 0
  mvarValorFinal = 0
  mvarTipoDeMensaje = 48
  mvarGrupo = "General"
  mvarProceder = True
  mvarExigirDatos = True
  mvarOmitirBlancosEnExtremos = True
End Sub
