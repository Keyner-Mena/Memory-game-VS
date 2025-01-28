Public Class Form1
    Public a As Integer ' Número aleatorio
    Const TamaTab As Byte = 4 ' Dimensión del tablero
    Public contadorElementos As Dictionary(Of Integer, Integer) ' Contador de valores
    Public primerClick As Boolean
    Public oprimioPic As String
    Public ultimoPicPress As String
    Public sumaPuntos As Integer = 0
    Public segundos As Integer = 0
    Public minutos As Integer = 0
    Public intentos As Integer = 0
    Public pictureBoxes As PictureBox()

    Public Sub ValoresGlobales()
        contadorElementos = New Dictionary(Of Integer, Integer)
        sumaPuntos = 0
        ultimoPicPress = ""
        oprimioPic = ""
        minutos = 0
        segundos = 0
        intentos = 0
        lblminutos.Text = "00"
        TextBox1.Text = $"{sumaPuntos} /16"
    End Sub

    Public Function yaEsta2Veces(valor As Integer) As Boolean
        ' Verificar si el valor ya está dos veces en el diccionario
        If contadorElementos.ContainsKey(valor) AndAlso contadorElementos(valor) >= 2 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub AsignarImagenAlPictureBox(picBox As PictureBox, valor As Integer)
        picBox.Tag = Trim(Str(valor))
        picBox.Image = Image.FromFile("Parte_Atras.png")
        picBox.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

    Private Sub IncrementarContador(valor As Integer)
        ' Incrementar el contador de un valor en el diccionario
        If contadorElementos.ContainsKey(valor) Then
            contadorElementos(valor) += 1
        Else
            contadorElementos(valor) = 1
        End If
    End Sub

    Public Sub darVuelta(ultimoPicPressPar As String)
        Dim index As Integer = Val(ultimoPicPressPar) - 1
        If index >= 0 AndAlso index < pictureBoxes.Length Then
            pictureBoxes(index).Image = Image.FromFile("Parte_Atras.png")
        End If
    End Sub

    Public Sub desHabilitar(ultimoPicPressPar As String)
        Dim index As Integer = Val(ultimoPicPressPar) - 1
        If index >= 0 AndAlso index < pictureBoxes.Length Then
            pictureBoxes(index).Enabled = False
            pictureBoxes(index).Tag = "0"
        End If
    End Sub

    Public Sub deshabiliteExcepto(uno As Integer, dos As Integer)
        For Each pic In pictureBoxes
            pic.Enabled = False
        Next

        If uno > 0 AndAlso uno <= pictureBoxes.Length Then
            pictureBoxes(uno - 1).Enabled = True
        End If

        If dos > 0 AndAlso dos <= pictureBoxes.Length Then
            pictureBoxes(dos - 1).Enabled = True
        End If
    End Sub

    Public Sub HabilitarTodosBotones()
        For Each picBox As PictureBox In pictureBoxes
            picBox.Enabled = True
        Next
    End Sub

    Public Sub preparativosIniciales()
        On Error GoTo ErrorInicial
        primerClick = True
        Randomize()

        For k = 1 To TamaTab
            For p = 1 To TamaTab
                Dim Encontro As Boolean = False
                Do While Not Encontro
                    a = Int(Rnd() * (TamaTab * 2)) + 1
                    If Not yaEsta2Veces(a) Then
                        ' Asignar el valor al PictureBox correspondiente
                        Dim picIndex = (k - 1) * TamaTab + p
                        If picIndex >= 1 AndAlso picIndex <= 16 Then
                            AsignarImagenAlPictureBox(pictureBoxes(picIndex - 1), a)
                        End If
                        IncrementarContador(a)
                        Encontro = True
                    End If
                Loop
            Next p
        Next k
        Exit Sub

        ErrorInicial:
        MsgBox("Error Falta Archivo: " & Err.Description)
        Resume Next
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Inicializar la colección de PictureBoxes
        pictureBoxes = {PictureBox1, PictureBox2, PictureBox3, PictureBox4,
                        PictureBox5, PictureBox6, PictureBox7, PictureBox8,
                        PictureBox9, PictureBox10, PictureBox11, PictureBox12,
                        PictureBox13, PictureBox14, PictureBox15, PictureBox16}

        ' Deshabilitar botones y establecer controladores de eventos
        For Each picBox As PictureBox In pictureBoxes
            picBox.Enabled = False
            AddHandler picBox.Click, AddressOf PictureBox_Click
        Next

        ' Preparativos iniciales
        ValoresGlobales()
        preparativosIniciales()

        ' Configurar estado inicial de botones
        Me.Button1.Enabled = False
    End Sub

    Private Sub PictureBox_Click(sender As Object, e As EventArgs)
        Dim picBox As PictureBox = CType(sender, PictureBox)
        Dim picBoxNumber As String = picBox.Name.Replace("PictureBox", "")

        ' Evitar que se pueda seleccionar el mismo PictureBox dos veces seguidas
        If Not primerClick AndAlso picBoxNumber = ultimoPicPress Then
            MsgBox("¡No puedes seleccionar el mismo cuadro dos veces!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        picBox.Image = Image.FromFile(picBox.Tag & ".png")

        If primerClick Then
            primerClick = False
            oprimioPic = Trim(picBox.Tag)
            ultimoPicPress = picBoxNumber
            intentos += 1
        Else
            primerClick = True
            If Val(picBox.Tag) = Val(oprimioPic) Then
                picBox.Enabled = False
                picBox.Tag = "0"
                desHabilitar(ultimoPicPress)
                sumaPuntos += 2
                TextBox1.Text = $"{sumaPuntos} /16"

                If sumaPuntos >= 16 Then
                    Timer1.Stop()
                    MsgBox($"¡Felicidades, has ganado! Tiempo Total: {Trim(lblminutos.Text & " :" & Label4.Text)} - Intentos: {intentos}")
                End If
            Else
                deshabiliteExcepto(CInt(picBoxNumber), Val(ultimoPicPress))
                For k = 1 To 600000
                    Application.DoEvents()
                Next
                HabilitarTodosBotones()
                picBox.Image = Image.FromFile("Parte_Atras.png")
                darVuelta(ultimoPicPress)
            End If
        End If
    End Sub

    ' Boton de reinicio
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ValoresGlobales()
        HabilitarTodosBotones()
        preparativosIniciales()
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        segundos += 1
        Label4.Text = If(segundos > 9, Str(segundos), "0" & Str(segundos))

        If segundos = 59 Then
            segundos = 0
            minutos += 1
            lblminutos.Text = "0" & Str(minutos)
        End If
    End Sub

    Private Sub btnIniciar_Click(sender As Object, e As EventArgs) Handles btnIniciar.Click
        MsgBox("El temporizador va a comenzar! Arma parejas de cartas para ganar", MsgBoxStyle.Information)

        Timer1.Enabled = True
        Timer1.Start()

        Me.Button1.Enabled = True
        btnIniciar.Enabled = False

        For Each picBox As PictureBox In pictureBoxes
            picBox.Enabled = True
        Next
    End Sub
End Class
