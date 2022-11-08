VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarcaMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Mensaje"
'---------------------------------------------------------------------------------------
' Module    : MarcaMensaje
' DateTime  : 27/08/2003 20:01
' Author    : adgarza
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Public Key As String

Private mvarCualControl As Object  'copia local
Private mvarMensaje As String 'copia local
Private mvarMensajeCompuesto As String 'copia local
Private mvarTag As String 'copia local
' Private WithEvents imgImagen As Image  ' El Image generado
Private WithEvents pctImagen As PictureBox   ' El Image generado
Attribute pctImagen.VB_VarHelpID = -1
Private frmFormulario As Object  ' El formulario principal
Private ctlContenedor As Object  ' El contenedor
Private mvarTipoDeMensaje As TipoMensaje 'copia local
Private strNombreControl As String

'---------------------------------------------------------------------------------------
' Procedure : TipoDeMensaje
' DateTime  : 27/08/2003 20:01
' Author    : adgarza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let TipoDeMensaje(ByVal vData As TipoMensaje)
    mvarTipoDeMensaje = vData
End Property

Public Property Get TipoDeMensaje() As TipoMensaje
    TipoDeMensaje = mvarTipoDeMensaje
End Property

Public Property Let Tag(ByVal vData As String)
    mvarTag = vData
End Property

Public Property Get Tag() As String
    Tag = mvarTag
End Property

Public Property Let MensajeCompuesto(ByVal vData As String)
    mvarMensajeCompuesto = vData
End Property

Public Property Get MensajeCompuesto() As String
    MensajeCompuesto = mvarMensajeCompuesto
End Property

Public Property Let Mensaje(ByVal vData As String)
    mvarMensaje = vData
End Property

Public Property Get Mensaje() As String
    Mensaje = mvarMensaje
End Property

Public Sub Mostrar()
  If Not CualControl Is Nothing Then
    If Not MostrarIcono Is Nothing Then
      'imgImagen.BorderStyle = vbBSNone
      'Set imgImagen = MostrarIcono
      pctImagen.BackColor = ctlContenedor.BackColor
      pctImagen.AutoRedraw = True
      pctImagen.PaintPicture MostrarIcono, 0, 0, frmFormulario.ScaleX(16, vbPixels, vbTwips), frmFormulario.ScaleY(16, vbPixels, vbTwips)
      pctImagen.AutoRedraw = False
    Else
      ' Establece las propiedades de acuerdo con el tipo de mensaje por mostrar
      Select Case mvarTipoDeMensaje
        Case Informacion
          pctImagen.BackColor = vbBlue
        Case Critico
          pctImagen.BackColor = vbRed
        Case Exclamacion
          pctImagen.BackColor = vbYellow
        Case Pregunta
          pctImagen.BackColor = vbGreen
      End Select
    End If
    
    ' imgSimbolo.Visible = True
    pctImagen.ToolTipText = Mensaje
    pctImagen.Visible = True
    pctImagen.ZOrder
  End If
End Sub

Public Property Set CualControl(ByVal vData As Object)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.CualControl = Form1
    Set mvarCualControl = vData
    
    ' Establece los controles
    Set frmFormulario = CualForm
    Set ctlContenedor = CualContenedor
    ' Set imgImagen = CualImage
    Set pctImagen = CualPicture
    
    ' Establece las propiedades y acomoda el cuadro de imagen
    'imgImagen.Stretch = True
    'imgImagen.Move CualControl.Left + CualControl.Width + 30, CualControl.Top, frmFormulario.ScaleX(16, vbPixels, vbTwips), frmFormulario.ScaleY(16, vbPixels, vbTwips)
    'imgImagen.ToolTipText = Mensaje
    
    ' Establece las propiedades y acomoda el cuadro de dibujo
    pctImagen.Appearance = 0
    pctImagen.BackColor = ctlContenedor.BackColor
    pctImagen.BorderStyle = vbBSNone
    pctImagen.Move CualControl.Left + CualControl.Width + 30, CualControl.Top, frmFormulario.ScaleX(16, vbPixels, vbTwips), frmFormulario.ScaleY(16, vbPixels, vbTwips)
    pctImagen.ToolTipText = Mensaje
    
    
End Property

Public Property Get CualControl() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CualControl
    Set CualControl = mvarCualControl
End Property

Private Property Get CualForm() As Form
  ' Var
  Dim objFrm As Object
  
  Set objFrm = CualControl
  Do
    Set objFrm = objFrm.Container
  Loop Until TypeOf objFrm Is Form
  
  Set CualForm = objFrm
End Property

Private Property Get CualContenedor() As Object
  Set CualContenedor = CualControl.Container
End Property

Private Property Get MostrarIcono() As StdPicture
  ' Var
  Dim strIcono As String
  Dim img As StdPicture
  
  Select Case mvarTipoDeMensaje
    Case Critico
      strIcono = "icrit.ico"
    Case Pregunta
      strIcono = "ipreg.ico"
    Case Informacion
      strIcono = "iinfo.ico"
    Case Exclamacion
      strIcono = "iexcl.ico"
  End Select
  
  On Error Resume Next
  Set img = LoadPicture(App.Path & "\" & strIcono, vbLPSmallShell)
  If Err.Number <> 0 Then
    Set img = Nothing
    Err.Clear
  End If
  On Error GoTo 0
  Set MostrarIcono = img
End Property

Private Property Get CualImage() As Image
  ' Var
  Dim img As Image
  Dim i As Integer
  
  On Error Resume Next
  Do
    i = i + 1
    Set img = frmFormulario.Controls.Add("VB.Image", "imgMensaje" & i, ctlContenedor)
  Loop Until Not img Is Nothing
  On Error GoTo 0
  Set CualImage = img
End Property

Private Sub imgImagen_Click()
  MsgBox MensajeCompuesto, TipoDeMensaje, App.Title
End Sub

Private Property Get CualPicture() As PictureBox
  ' Var
  Dim pct As PictureBox
  Dim i As Integer
  
  On Error Resume Next
  Do
    i = i + 1
    strNombreControl = "pctMensaje" & i
    Set pct = frmFormulario.Controls.Add("VB.PictureBox", strNombreControl, ctlContenedor)
  Loop Until Not pct Is Nothing
  On Error GoTo 0
  Set CualPicture = pct
End Property

Private Sub Class_Terminate()
  On Error Resume Next
    pctImagen.Visible = False
    frmFormulario.Controls.Remove strNombreControl
  On Error GoTo 0
  
  Set pctImagen = Nothing
  Set ctlContenedor = Nothing
  Set frmFormulario = Nothing
  Set mvarCualControl = Nothing
End Sub

Private Sub pctImagen_Click()
  On Error Resume Next
  CualControl.SetFocus
  CualControl.SelStart = 0
  CualControl.SelLength = Len(CualControl.Text)
  On Error GoTo 0
End Sub

Private Sub pctImagen_DblClick()
  MsgBox MensajeCompuesto, TipoDeMensaje, App.Title
End Sub
