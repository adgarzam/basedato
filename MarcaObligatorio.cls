VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarcaObligatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarCualControl As Object  'copia local
Private mvarMensaje As String 'copia local
Private mvarMensajeCompuesto As String 'copia local
Private mvarTag As String 'copia local
' Private WithEvents imgImagen As Image  ' El Image generado
Private WithEvents pctObligatorio As PictureBox   ' El Image generado
Attribute pctObligatorio.VB_VarHelpID = -1
Private frmFormulario As Object  ' El formulario principal
Private ctlContenedor As Object  ' El contenedor
Private mvarTipoDeMarcado As TipoMarcado 'copia local
Private strNombreControl As String

Public Property Let TipoDeMarcado(ByVal vData As TipoMarcado)
    mvarTipoDeMarcado = vData
End Property


Public Property Get TipoDeMarcado() As TipoMarcado
    TipoDeMarcado = mvarTipoDeMarcado
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
      pctObligatorio.BackColor = ctlContenedor.BackColor
      pctObligatorio.AutoRedraw = True
      pctObligatorio.PaintPicture MostrarIcono, 0, 0, frmFormulario.ScaleX(16, vbPixels, vbTwips), frmFormulario.ScaleY(16, vbPixels, vbTwips)
      pctObligatorio.AutoRedraw = False
    Else
      ' Establece las propiedades de acuerdo con el tipo de mensaje por mostrar
      Select Case mvarTipoDeMarcado
        Case Obligatorio
          pctObligatorio.BackColor = vbBlack
'        Case Optativo
'          pctObligatorio.BackColor = vbWhite
'        Case Formato
'          pctObligatorio.BackColor = QBColor(8)
      End Select
    End If
    
    ' imgSimbolo.Visible = True
    pctObligatorio.ToolTipText = Mensaje
    pctObligatorio.Visible = True
  End If
End Sub

Public Property Set CualControl(ByVal vData As Object)
    Set mvarCualControl = vData
    
    ' Establece los controles
    Set frmFormulario = CualForm
    Set ctlContenedor = CualContenedor
    ' Set imgImagen = CualImage
    Set pctObligatorio = CualPicture
    
    ' Establece las propiedades y acomoda el cuadro de imagen
    'imgImagen.Stretch = True
    'imgImagen.Move CualControl.Left + CualControl.Width + 30, CualControl.Top, frmFormulario.ScaleX(16, vbPixels, vbTwips), frmFormulario.ScaleY(16, vbPixels, vbTwips)
    'imgImagen.ToolTipText = Mensaje
    
    ' Establece las propiedades y acomoda el cuadro de dibujo
    pctObligatorio.Appearance = 0
    pctObligatorio.BackColor = ctlContenedor.BackColor
    pctObligatorio.BorderStyle = vbBSNone
    pctObligatorio.Move CualControl.Left + CualControl.Width + 30, CualControl.Top, frmFormulario.ScaleX(16, vbPixels, vbTwips), frmFormulario.ScaleY(16, vbPixels, vbTwips)
    pctObligatorio.ToolTipText = Mensaje
        
End Property

Public Property Get CualControl() As Object
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
  
  Select Case mvarTipoDeMarcado
    Case Obligatorio
      strIcono = "iobli.ico"
'    Case Optativo
'      strIcono = "iopta.ico"
'    Case Formato
'      strIcono = "iform.ico"
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

Private Property Get CualPicture() As PictureBox
  ' Var
  Dim pct As PictureBox
  Dim i As Integer
  
  On Error Resume Next
  Do
    i = i + 1
    strNombreControl = "pctObligatorio" & i
    Set pct = frmFormulario.Controls.Add("VB.PictureBox", strNombreControl, ctlContenedor)
  Loop Until Not pct Is Nothing
  On Error GoTo 0
  Set CualPicture = pct
End Property

Private Sub Class_Initialize()
  mvarMensaje = "Este campo es obligatorio."
  mvarMensajeCompuesto = "Debe teclear el valor correspondiente a este campo. Su valor es necesario para el correcto funcionamiento del programa."
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
    pctObligatorio.Visible = False
    frmFormulario.Controls.Remove strNombreControl
  On Error GoTo 0
  
  Set pctObligatorio = Nothing
  Set ctlContenedor = Nothing
  Set frmFormulario = Nothing
  Set mvarCualControl = Nothing
End Sub

Private Sub pctObligatorio_Click()
  On Error Resume Next
  CualControl.SetFocus
  CualControl.SelStart = 0
  CualControl.SelLength = Len(CualControl.Text)
  On Error GoTo 0
End Sub

Private Sub pctObligatorio_DblClick()
  MsgBox MensajeCompuesto, vbInformation, App.Title
End Sub

