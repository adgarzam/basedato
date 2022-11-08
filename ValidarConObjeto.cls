VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidarConObjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Mensaje"
Option Explicit

Public Key As String

'variables locales para almacenar los valores de las propiedades
Private mvarControlPorValidar As Object 'copia local
Private mvarControlParaComparacion As Object 'copia local
Private mvarMensaje As String 'copia local
Private mvarNombreColoquial As String 'copia local
Private mvarNombreColoquialControlParaComparacion As String 'copia local
Private mvarComparacion As TipoEvaluacion 'copia local
Private mvarValidado As Boolean 'copia local
Private mvarTipoDeMensaje As TipoMensaje 'copia local
Private mvarGrupo As String 'copia local
Private mvarTipoDeValores As TipoValores 'copia local
Private mvarEsValido As Boolean 'copia local
Private strValorAnterior1 As String
Private strValorAnterior2 As String
Private mvarProceder As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarExigirDatos As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
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

Public Property Let EsValido(ByVal vData As Boolean)
    mvarEsValido = vData
End Property

Public Property Get EsValido() As Boolean
   If Not Me.Validado Then
     Me.Validar
   End If
   EsValido = mvarEsValido
End Property

Public Property Let TipoDeValores(ByVal vData As TipoValores)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.TipoDeValores = Form1
    mvarTipoDeValores = vData
End Property

Public Property Get TipoDeValores() As TipoValores
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.TipoDeValores
    TipoDeValores = mvarTipoDeValores
End Property

Public Property Let Grupo(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Grupo = 5
    mvarGrupo = vData
End Property

Public Property Get Grupo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Grupo
    Grupo = mvarGrupo
End Property

Public Property Let TipoDeMensaje(ByVal vData As TipoMensaje)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.TipoDeMensaje = 5
    mvarTipoDeMensaje = vData
End Property

Public Property Get TipoDeMensaje() As TipoMensaje
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.TipoDeMensaje
    TipoDeMensaje = mvarTipoDeMensaje
End Property

Public Function Validar() As Boolean
  ' Var
  Dim bolComparar As Boolean
  Dim bolValidar As Boolean
  Dim varValor1 As Variant
  Dim varValor2 As Variant
  
  strValorAnterior1 = Me.ControlPorValidar.Text
  strValorAnterior2 = Me.ControlParaComparacion.Text
  
  If Me.OmitirBlancosEnExtremos Then
    strValorAnterior1 = Trim(strValorAnterior1)
    strValorAnterior2 = Trim(strValorAnterior2)
  End If
  
  If Len(strValorAnterior1) > 0 And Len(strValorAnterior2) > 0 Then
    Select Case Me.TipoDeValores
      Case TipoValores.General
        varValor1 = strValorAnterior1
        varValor2 = strValorAnterior2
        bolComparar = True
      Case TipoValores.Cadena
        varValor1 = strValorAnterior1
        varValor2 = strValorAnterior2
        bolComparar = True
      Case TipoValores.CadenaIgnorarMayYMin
        varValor1 = UCase(strValorAnterior1)
        varValor2 = UCase(strValorAnterior2)
        bolComparar = True
      Case TipoValores.Numeros
        If IsNumeric(strValorAnterior1) And IsNumeric(strValorAnterior2) Then
          varValor1 = Val(strValorAnterior1)
          varValor2 = Val(strValorAnterior2)
          bolComparar = True
        Else
          bolComparar = False
        End If
      Case TipoValores.Fechas
        If IsDate(strValorAnterior1) And IsDate(strValorAnterior2) Then
          varValor1 = CDate(strValorAnterior1)
          varValor2 = CDate(strValorAnterior2)
          bolComparar = True
        Else
          bolComparar = False
        End If
    End Select
    
    If bolComparar Then
      Select Case Me.Comparacion
        Case TipoEvaluacion.Distinto
          bolValidar = (varValor1 <> varValor2)
        Case TipoEvaluacion.Igual
          bolValidar = (varValor1 = varValor2)
        Case TipoEvaluacion.MayorIgual
          bolValidar = (varValor1 >= varValor2)
        Case TipoEvaluacion.MayorQue
          bolValidar = (varValor1 > varValor2)
        Case TipoEvaluacion.MenorIgual
          bolValidar = (varValor1 <= varValor2)
        Case TipoEvaluacion.MenorQue
          bolValidar = (varValor1 < varValor2)
      End Select
    Else
      bolValidar = False
    End If
  Else
    Me.EsValido = Not Me.ExigirDatos
  End If
      
  Me.EsValido = bolValidar
  Me.Validado = True
  Validar = Me.EsValido

End Function

Public Property Let Validado(ByVal vData As Boolean)
    mvarValidado = vData
End Property


Public Property Get Validado() As Boolean
  If Me.OmitirBlancosEnExtremos Then
    Me.Validado = (strValorAnterior1 = Trim(Me.ControlPorValidar.Text)) And (strValorAnterior2 = Trim(Me.ControlParaComparacion.Text))
  Else
    Me.Validado = (strValorAnterior1 = Me.ControlPorValidar.Text) And (strValorAnterior2 = Me.ControlParaComparacion.Text)
  End If
  Validado = mvarValidado
End Property

Public Property Let Comparacion(ByVal vData As TipoEvaluacion)
    If vData <> CumpleConPatron Then
        mvarComparacion = vData
    End If
End Property

Public Property Get Comparacion() As TipoEvaluacion
    Comparacion = mvarComparacion
End Property

Public Property Let NombreColoquialControlParaComparacion(ByVal vData As String)
    mvarNombreColoquialControlParaComparacion = vData
End Property

Public Property Get NombreColoquialControlParaComparacion() As String
    NombreColoquialControlParaComparacion = mvarNombreColoquialControlParaComparacion
End Property

Public Property Let NombreColoquial(ByVal vData As String)
    mvarNombreColoquial = vData
End Property


Public Property Get NombreColoquial() As String
    NombreColoquial = mvarNombreColoquial
End Property

Public Property Let Mensaje(ByVal vData As String)
    mvarMensaje = vData
End Property

Public Property Get Mensaje() As String
    Mensaje = mvarMensaje
End Property

Public Property Set ControlParaComparacion(ByVal vData As Object)
     ' Var
    Dim strTipo As String
    
   If mvarControlParaComparacion Is Nothing Then
      On Error Resume Next
      strTipo = vData.Text
      If Err.Number = 0 Then
        On Error GoTo 0
        Set mvarControlParaComparacion = vData
      Else
        MsgBox "Este validador sólo admite controles que cuenten con una propiedad Text.", vbCritical, "Error " & Err.Number
        Err.Clear
        On Error GoTo 0
      End If
   Else
     MsgBox "Sólo se puede asignar una vez el control.", vbExclamation, "Error"
   End If
End Property

Public Property Get ControlParaComparacion() As Object
  Set ControlParaComparacion = mvarControlParaComparacion
End Property

Public Property Set ControlPorValidar(ByVal vData As Object)
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
  mvarMensaje = "La comparación entre objetos no fue satisfactoria."
  mvarTipoDeMensaje = 48
  mvarGrupo = "General"
  mvarComparacion = Igual
  mvarTipoDeValores = TipoValores.General
  mvarProceder = True
  mvarExigirDatos = True
  mvarOmitirBlancosEnExtremos = True
End Sub
