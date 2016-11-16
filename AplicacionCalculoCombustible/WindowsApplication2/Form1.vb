REM Curso Visual Basic aprenderaprogramar.com
Option Explicit On
Public Class Form1
    REM Declaración de variables
    Dim camionetas, autos, motos As Integer
    Dim Capcamion, Capauto, Capmoto As Single
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        limpiarCampos(Me)
    End Sub

    Dim Necesidadescom As Single

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    REM Contenido del formulario
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Cálculo de necesidades combustible"
        Label1.Text = "Favor ingresar datos"
        Label2.Text = "Camionetas"
        Label3.Text = "Autos"
        Label7.Text = "Motos"
        Label4.Text = "Cantidad de Equipos (Unidades)"
        Label5.Text = "Capacidad de Equipos (Galones)"

        'Button1.Text = "Aceptar"
    End Sub

    REM Cálculo y muestra resultados
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Len(TextBox1.Text) = 0 Then
            Label8.Text = ("DATO 1  ES INVALIDO")
        ElseIf Len(TextBox2.Text) = 0 Then
            Label8.Text = ("DATO 2 ES INVALIDO")
        ElseIf Len(TextBox3.Text) = 0 Then
            Label8.Text = ("DATO 3 ES INVALIDO")
        ElseIf Len(TextBox4.Text) = 0 Then
            Label8.Text = ("DATO 4 ES INVALIDO")
        ElseIf Len(TextBox5.Text) = 0 Then
            Label8.Text = ("DATO 5 ES INVALIDO")
        ElseIf Len(TextBox6.Text) = 0 Then
            Label8.Text = ("DATO 6 ES INVALIDO")
        End If
        Label6.ForeColor = Color.Black
        Label6.Font = New Font("Arial", 11, FontStyle.Bold)
        camionetas = Val(TextBox1.Text)
        autos = Val(TextBox2.Text)
        motos = Val(TextBox5.Text)
        Capcamion = Val(TextBox3.Text)
        Capauto = Val(TextBox4.Text)
        Capmoto = Val(TextBox6.Text)
        Necesidadescom = camionetas * Capcamion + autos * Capauto + motos * Capmoto
        Label6.Text = "El requerimiento de Combustible para hoy es " & Necesidadescom & " Galones"
    End Sub
End Class