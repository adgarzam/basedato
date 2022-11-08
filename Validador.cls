VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Validador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MarcaMensajes"
Attribute VB_Ext_KEY = "Member1" ,"MarcaObligatorios"
Attribute VB_Ext_KEY = "Member2" ,"ValidarConObjetos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member3" ,"ValidarExpresiones"
Attribute VB_Ext_KEY = "Member4" ,"ValidarFechas"
Attribute VB_Ext_KEY = "Member5" ,"ValidarNumericos"
Attribute VB_Ext_KEY = "Member6" ,"ValidarRangos"
Attribute VB_Ext_KEY = "Member7" ,"ValidarVacios"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarValidarVacios As ValidarVacios
Private mvarValidarExpresiones As ValidarExpresiones
Private mvarValidarRangos As ValidarRangos
Private mvarEsValido As Boolean 'copia local
Private mvarValidarConObjetos As ValidarConObjetos
Private mvarValidarNumericos As ValidarNumericos
Private mvarMarcaMensajes As MarcaMensajes 'copia local
Private mvarValidarFechas As ValidarFechas ' copia local
Private mvarMarcaObligatorios As MarcaObligatorios ' copia local
Private mvarFormularioPorValidar As Form 'copia local
Private bolAsignadoPorPropiedad As Boolean

' Enumeraciones
Public Enum TipoMarcado
    Obligatorio = 1
'     Optativo = 2
'     Formato = 4
'     General = 8
End Enum

Public Enum TipoMensaje
    Critico = 16
    Pregunta = 32
    Exclamacion = 48
    Informacion = 64
End Enum

Public Enum TipoEvaluacion
    MayorQue = 1
    MenorQue = 2
    Igual = 3
    MayorIgual = 4
    MenorIgual = 5
    Distinto = 6
    CumpleConPatron = 7
End Enum

Public Enum TipoNumero
    Entero = 1
    Real = 2
    Indistinto = 3
End Enum

Public Enum TipoValores
    General = 0
    Cadena = 1
    CadenaIgnorarMayYMin = 2
    Numeros = 3
    Fechas = 4
End Enum
Public Enum TipoValidacion
    Vacios
    Expresiones
    ConObjetos
    Fechas
    Numericos
    Rangos
End Enum

Public Sub LimpiarValidaciones()
  Me.ValidarVacios.Clear
  Me.ValidarConObjetos.Clear
  Me.ValidarNumericos.Clear
  Me.ValidarExpresiones.Clear
  Me.ValidarRangos.Clear
  Me.ValidarFechas.Clear
End Sub

Public Sub EstablecerValidacion(Optional Vacios As Boolean = True, Optional ConObjetos As Boolean = True, Optional Numericos As Boolean = True, Optional Expresiones As Boolean = True, Optional Rangos As Boolean = True, Optional Fechas As Boolean = True)
  Me.ValidarVacios.Proceder = Vacios
  Me.ValidarConObjetos.Proceder = ConObjetos
  Me.ValidarNumericos.Proceder = Numericos
  Me.ValidarExpresiones.Proceder = Expresiones
  Me.ValidarRangos.Proceder = Rangos
  Me.ValidarFechas.Proceder = Fechas
End Sub

Public Property Set FormularioPorValidar(ByVal vData As Object)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.FormularioPorValidar = Form1
    If mvarFormularioPorValidar Is Nothing Then
      Set mvarFormularioPorValidar = vData
      bolAsignadoPorPropiedad = True
      doEstablecerCronometro
    ElseIf vData Is mvarFormularioPorValidar Then
      ' Está correcto...
    Else
      If bolAsignadoPorPropiedad Then
        MsgBox "No se puede asignar un formulario distinto al ya asignado.", vbExclamation, "Error"
      Else
        MsgBox "Ya se tenía un formulario reconocido en la asignación de controles. No se puede cambiar.", vbExclamation, "Error"
      End If
    End If
End Property


Public Property Get FormularioPorValidar() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FormularioPorValidar
    Set FormularioPorValidar = mvarFormularioPorValidar
End Property

Public Property Get MarcaObligatorios() As MarcaObligatorios
    If mvarMarcaObligatorios Is Nothing Then
        Set mvarMarcaObligatorios = New MarcaObligatorios
    End If

    Set MarcaObligatorios = mvarMarcaObligatorios
End Property

Public Property Set MarcaObligatorios(vData As MarcaObligatorios)
    Set mvarMarcaObligatorios = vData
End Property

Public Property Get ValidarFechas() As ValidarFechas
    If mvarValidarFechas Is Nothing Then
        Set mvarValidarFechas = New ValidarFechas
    End If
    
    Set ValidarFechas = mvarValidarFechas
End Property

Public Property Set ValidarFechas(vData As ValidarFechas)
    Set mvarValidarFechas = vData
End Property

Public Property Set MarcaMensajes(ByVal vData As MarcaMensajes)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.MarcaMensajes = Form1
    Set mvarMarcaMensajes = vData
End Property

Public Property Get MarcaMensajes() As MarcaMensajes
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.MarcaMensajes
    If mvarMarcaMensajes Is Nothing Then
        Set mvarMarcaMensajes = New MarcaMensajes
    End If
    
    Set MarcaMensajes = mvarMarcaMensajes
End Property

Public Sub MarcarMensajes()
  MarcaMensajes.Mostrar
End Sub

Public Sub MarcarObligatorios()
  ' Var
  Dim vva As ValidarVacio
  Dim mob As MarcaObligatorio
  
  ' Primero, verificamos cuáles son los obligatorios
  With MarcaObligatorios
    .Clear
    For Each vva In ValidarVacios
      Set mob = .Add(vva.ControlPorValidar)
    Next vva
    
    .Mostrar
  End With
  
End Sub

Public Property Get ValidarNumericos() As ValidarNumericos
    If mvarValidarNumericos Is Nothing Then
        Set mvarValidarNumericos = New ValidarNumericos
    End If

    Set ValidarNumericos = mvarValidarNumericos
End Property


Public Property Set ValidarNumericos(vData As ValidarNumericos)
    Set mvarValidarNumericos = vData
End Property

Public Property Get ValidarConObjetos() As ValidarConObjetos
    If mvarValidarConObjetos Is Nothing Then
        Set mvarValidarConObjetos = New ValidarConObjetos
    End If

    Set ValidarConObjetos = mvarValidarConObjetos
End Property

Public Property Set ValidarConObjetos(vData As ValidarConObjetos)
    Set mvarValidarConObjetos = vData
End Property

Public Property Let EsValido(ByVal vData As Boolean)
    mvarEsValido = vData
End Property

Public Property Get EsValido() As Boolean
  ' Var
  Dim vva As ValidarVacio
  Dim vra As ValidarRango
  Dim vex As ValidarExpresion
  Dim vco As ValidarConObjeto
  Dim vfe As ValidarFecha
  Dim bolRealizarValidacion As Boolean
  
  bolRealizarValidacion = False
  
  ' Verifica si todo se ha validado ya
  If ValidarVacios.Proceder Then
    If ValidarVacios.Count > 0 Then
      For Each vva In ValidarVacios
        If vva.Validado = False And vva.Proceder = True Then
          bolRealizarValidacion = True
        End If
      Next vva
    End If
  End If
  
  ' Valida si han quedado campos fuera de rango
  If ValidarRangos.Proceder Then
    If ValidarRangos.Count > 0 Then
      For Each vra In ValidarRangos
        If vra.Validado = False And vra.Proceder = True Then
          bolRealizarValidacion = True
        End If
      Next vra
    End If
  End If
  
  ' Valida si hay campos que no cumplan con la expresión
  If ValidarExpresiones.Proceder Then
    If ValidarExpresiones.Count > 0 Then
      For Each vex In ValidarExpresiones
        If vex.Validado = False And vex.Proceder = True Then
          bolRealizarValidacion = True
        End If
      Next vex
    End If
  End If

  ' Valida si la comparación con otro objeto no confluye
  If ValidarConObjetos.Proceder Then
    If ValidarConObjetos.Count > 0 Then
      For Each vco In ValidarConObjetos
        If vco.Validado = False And vco.Proceder = True Then
          bolRealizarValidacion = True
        End If
      Next vco
    End If
  End If
    
  ' Valida si la validación de una fecha no congfluye
  If ValidarFechas.Proceder Then
    If ValidarFechas.Count > 0 Then
      For Each vfe In ValidarFechas
        If vfe.Validado = False And vfe.Proceder = True Then
          bolRealizarValidacion = True
        End If
      Next vfe
    End If
  End If
    
  If bolRealizarValidacion Then
    Me.Validar
  End If
  
  EsValido = mvarEsValido
End Property

Public Function Validar(Optional CualGrupo As String = "", Optional IncluirGeneral As Boolean = True) As Boolean
  ' Var
  Dim vva As ValidarVacio
  Dim vfe As ValidarFecha
  Dim vra As ValidarRango
  Dim vex As ValidarExpresion
  Dim vco As ValidarConObjeto
  Dim vnu As ValidarNumerico
  Dim men As MarcaMensaje
  Dim bolValidar As Boolean
  
  ' Limpia los mensajes anteriores
  MarcaMensajes.Clear
  
  ' Valida si han quedado campos vacíos
  If ValidarVacios.Proceder Then
    If ValidarVacios.Count > 0 Then
      For Each vva In ValidarVacios
        If CualGrupo = "" Then
          bolValidar = True
        ElseIf UCase(CualGrupo) = UCase(vva.Grupo) Then
          bolValidar = True
        ElseIf IncluirGeneral And UCase(vva.Grupo = "GENERAL") Then
          bolValidar = True
        Else
          bolValidar = False
        End If
        
        If bolValidar Then
          If vva.Proceder Then
            If vva.Validar = False Then
              Set men = MarcaMensajes.Add(vva.ControlPorValidar, vva.Mensaje, vva.TipoDeMensaje)
              men.MensajeCompuesto = "El campo está vacío en " & vva.NombreColoquial & ". Este campo no puede quedar vacío. Verifique, por favor."
            End If
          End If
        End If
      Next vva
    End If
  End If
  
  ' Valida si han quedado campos que no sean fechas
  If ValidarFechas.Proceder Then
    If ValidarFechas.Count > 0 Then
      For Each vfe In ValidarFechas
        If CualGrupo = "" Then
          bolValidar = True
        ElseIf UCase(CualGrupo) = UCase(vfe.Grupo) Then
          bolValidar = True
        ElseIf IncluirGeneral And UCase(vfe.Grupo = "GENERAL") Then
          bolValidar = True
        Else
          bolValidar = False
        End If
        
        If bolValidar Then
          If vfe.Proceder Then
            If vfe.Validar = False Then
              Set men = MarcaMensajes.Add(vfe.ControlPorValidar, vfe.Mensaje, vfe.TipoDeMensaje)
              men.MensajeCompuesto = "El campo no tiene un valor de fecha en " & vfe.NombreColoquial & ". Teclee un formato de fecha adecuado, por favor."
            End If
          End If
        End If
      Next vfe
    End If
  End If
  
  
  ' Valida si han quedado campos fuera de rango
  If ValidarRangos.Proceder Then
    If ValidarRangos.Count > 0 Then
      For Each vra In ValidarRangos
        If CualGrupo = "" Then
          bolValidar = True
        ElseIf UCase(CualGrupo) = UCase(vra.Grupo) Then
          bolValidar = True
        ElseIf IncluirGeneral And UCase(vra.Grupo = "GENERAL") Then
          bolValidar = True
        Else
          bolValidar = False
        End If
        
        If bolValidar Then
          If vra.Proceder Then
            If vra.Validar = False Then
              Set men = MarcaMensajes.Add(vra.ControlPorValidar, vra.Mensaje, vra.TipoDeMensaje)
              men.MensajeCompuesto = "El valor está fuera de rango en " & vra.NombreColoquial & ". El rango debe estar entre " & vra.ValorInicial & " y " & vra.ValorFinal & "."
            End If
          End If
        End If
      Next vra
    End If
  End If
  
  ' Valida si hay campos que no cumplan con la expresión
  If ValidarExpresiones.Proceder Then
    If ValidarExpresiones.Count > 0 Then
      For Each vex In ValidarExpresiones
        If CualGrupo = "" Then
          bolValidar = True
        ElseIf UCase(CualGrupo) = UCase(vex.Grupo) Then
          bolValidar = True
        ElseIf IncluirGeneral And UCase(vex.Grupo = "GENERAL") Then
          bolValidar = True
        Else
          bolValidar = False
        End If
        
        If bolValidar Then
          If vex.Proceder Then
            If vex.Validar = False Then
              Set men = MarcaMensajes.Add(vex.ControlPorValidar, vex.Mensaje, vex.TipoDeMensaje)
              men.MensajeCompuesto = "El campo no cumple con la expresión en " & vex.NombreColoquial & ". " & vex.Mensaje & " Verifique, por favor."
            End If
          End If
        End If
      Next vex
    End If
  End If

  ' Valida si la comparación con otro objeto no confluye
  If ValidarConObjetos.Proceder Then
    If ValidarConObjetos.Count > 0 Then
      For Each vco In ValidarConObjetos
        If CualGrupo = "" Then
          bolValidar = True
        ElseIf UCase(CualGrupo) = UCase(vco.Grupo) Then
          bolValidar = True
        ElseIf IncluirGeneral And UCase(vco.Grupo = "GENERAL") Then
          bolValidar = True
        Else
          bolValidar = False
        End If
        
        If bolValidar Then
          If vco.Proceder Then
            If vco.Validar = False Then
              Set men = MarcaMensajes.Add(vco.ControlPorValidar, vco.Mensaje, vco.TipoDeMensaje)
              men.MensajeCompuesto = "Falló la comparación con el objeto en " & vco.NombreColoquial & ". " & vco.Mensaje & " Verifique, por favor."
            End If
          End If
        End If
      Next vco
    End If
  End If

  ' Valida si el valor es numérico
  If ValidarNumericos.Proceder Then
    If ValidarNumericos.Count > 0 Then
      For Each vnu In ValidarNumericos
        If CualGrupo = "" Then
          bolValidar = True
        ElseIf UCase(CualGrupo) = UCase(vnu.Grupo) Then
          bolValidar = True
        ElseIf IncluirGeneral And UCase(vnu.Grupo = "GENERAL") Then
          bolValidar = True
        Else
          bolValidar = False
        End If
        
        If bolValidar Then
          If vnu.Proceder Then
            If vnu.Validar = False Then
              Set men = MarcaMensajes.Add(vnu.ControlPorValidar, vnu.Mensaje, vnu.TipoDeMensaje)
              men.MensajeCompuesto = "Falló la comparación con el objeto en " & vnu.NombreColoquial & ". " & vnu.Mensaje & " Verifique, por favor."
            End If
          End If
        End If
      Next vnu
    End If
  End If

  ' EsValido final
  Me.EsValido = (MarcaMensajes.Count = 0)
  Validar = Me.EsValido
End Function

Public Property Get ValidarRangos() As ValidarRangos
    If mvarValidarRangos Is Nothing Then
        Set mvarValidarRangos = New ValidarRangos
    End If

    Set ValidarRangos = mvarValidarRangos
End Property

Public Property Set ValidarRangos(vData As ValidarRangos)
    Set mvarValidarRangos = vData
End Property

Public Property Get ValidarExpresiones() As ValidarExpresiones
    If mvarValidarExpresiones Is Nothing Then
        Set mvarValidarExpresiones = New ValidarExpresiones
    End If

    Set ValidarExpresiones = mvarValidarExpresiones
End Property


Public Property Set ValidarExpresiones(vData As ValidarExpresiones)
    Set mvarValidarExpresiones = vData
End Property

Public Property Get ValidarVacios() As ValidarVacios
    If mvarValidarVacios Is Nothing Then
        Set mvarValidarVacios = New ValidarVacios
    End If

    Set ValidarVacios = mvarValidarVacios
End Property

Public Property Set ValidarVacios(vData As ValidarVacios)
    Set mvarValidarVacios = vData
End Property

Private Sub Class_Terminate()
  Set mvarMarcaObligatorios = Nothing
  Set mvarValidarFechas = Nothing
  Set mvarValidarNumericos = Nothing
  Set mvarValidarConObjetos = Nothing
  Set mvarValidarRangos = Nothing
  Set mvarValidarExpresiones = Nothing
  Set mvarValidarVacios = Nothing
  Set mvarMarcaMensajes = Nothing
End Sub

Public Sub LimpiarTodo()
  ValidarVacios.Clear
  ValidarConObjetos.Clear
  ValidarExpresiones.Clear
  ValidarNumericos.Clear
  ValidarRangos.Clear
  MarcaMensajes.Clear
  MarcaObligatorios.Clear
End Sub

Private Sub doEstablecerCronometro()

End Sub
