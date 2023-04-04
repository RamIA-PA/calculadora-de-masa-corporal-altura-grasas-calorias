Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub











    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles apellido.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles nombre.TextChanged

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles ligero.CheckedChanged

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles moderado.CheckedChanged

    End Sub


    Function ICM()
        Dim masa As Double
        masa = (peso.Text / (altura.Text * 2))
        Return masa
    End Function



    Function ICE()
        Dim indicecintura As Double
        indicecintura = (cadera.Text / (altura.Text * 100))
        Return indicecintura
    End Function

    Function GC(masa)
        Dim grasa As Double
        If edad.Text <= 15 And masculino.Checked Then
            grasa = 1.51 * masa - 0.7 * edad.Text - 3.6 * 1 + 1.4
        ElseIf edad.Text <= 15 And femenino.Checked Then
            grasa = 1.51 * masa - 0.7 * edad.Text - 3.6 * 0 + 1.4
        ElseIf edad.Text > 15 And masculino.Checked Then
            grasa = 1.2 * masa + 0.23 * edad.Text - 10.8 * 1 - 5.4
        ElseIf edad.Text > 15 And femenino.Checked Then
            grasa = 1.2 * masa + 0.23 * edad.Text - 10.8 * 0 - 5.4
        End If
        Return grasa
    End Function

    Function CXD()
        Dim caldia As Double
        If femenino.Checked And sedentario.Checked Then
            caldia = 655 + (9.6 * peso.Text) + (1.8 * (altura.Text * 100)) - (4.7 * edad.Text) * 1.2
        ElseIf femenino.Checked And ligero.Checked Then
            caldia = 655 + (9.6 * peso.Text) + (1.8 * (altura.Text * 100)) - (4.7 * edad.Text) * 1.375
        ElseIf femenino.Checked And moderado.Checked Then
            caldia = 655 + (9.6 * peso.Text) + (1.8 * (altura.Text * 100)) - (4.7 * edad.Text) * 1.55
        ElseIf femenino.Checked And vigoroso.Checked Then
            caldia = 655 + (9.6 * peso.Text) + (1.8 * (altura.Text * 100)) - (4.7 * edad.Text) * 1.725
        ElseIf femenino.Checked And extremo.Checked Then
            caldia = 655 + (9.6 * peso.Text) + (1.8 * (altura.Text * 100)) - (4.7 * edad.Text) * 1.9

        ElseIf masculino.Checked And sedentario.Checked Then
            caldia = 66 + (13.7 * peso.Text) + (5 * (altura.Text * 100)) - (6.8 * edad.Text) * 1.2
        ElseIf masculino.Checked And ligero.Checked Then
            caldia = 66 + (13.7 * peso.Text) + (5 * (altura.Text * 100)) - (6.8 * edad.Text) * 1.375
        ElseIf masculino.Checked And moderado.Checked Then
            caldia = 66 + (13.7 * peso.Text) + (5 * (altura.Text * 100)) - (6.8 * edad.Text) * 1.55
        ElseIf masculino.Checked And vigoroso.Checked Then
            caldia = 66 + (13.7 * peso.Text) + (5 * (altura.Text * 100)) - (6.8 * edad.Text) * 1.725
        ElseIf masculino.Checked And extremo.Checked Then
            caldia = 66 + (13.7 * peso.Text) + (5 * (altura.Text * 100)) - (6.8 * edad.Text) * 1.9

        End If

        Return caldia
    End Function

    Function Leyendamasa()
        If masacor.Text < 19 Then

            leyenda.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "bajo peso"

        ElseIf masacor.Text >= 19 And masacor.Text <= 24.9 Then
            leyenda.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "peso normal"

        ElseIf masacor.Text >= 25 And masacor.Text <= 29.9 Then
            leyenda.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobrepeso"

        ElseIf masacor.Text > 30 Then
            leyenda.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "obesidad"
        End If
        Return leyenda.Text
    End Function


    Function Leyendaalcin()
        If indicintura.Text < 0.34 And (masculino.Checked And femenino.Checked) Then

            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "Delgadez severa"

        ElseIf indicintura.Text >= 0.35 And indicintura.Text <= 0.42 And masculino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "Delgadez leve"

        ElseIf indicintura.Text >= 0.35 And indicintura.Text <= 0.41 And femenino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "Delgadez leve"

        ElseIf indicintura.Text >= 0.43 And indicintura.Text <= 0.52 And masculino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "peso normal"

        ElseIf indicintura.Text >= 0.42 And indicintura.Text <= 0.48 And femenino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "peso normal"

        ElseIf indicintura.Text >= 0.53 And indicintura.Text <= 0.57 And masculino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobrepeso"

        ElseIf indicintura.Text >= 0.49 And indicintura.Text <= 0.53 And femenino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobrepeso"

        ElseIf indicintura.Text >= 0.58 And indicintura.Text <= 0.62 And masculino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobre peso elevedado"

        ElseIf indicintura.Text >= 0.54 And indicintura.Text <= 0.57 And femenino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobre peso elevedado"

        ElseIf indicintura.Text > 0.63 And masculino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "obesidad morbida"

        ElseIf indicintura.Text > 0.58 And femenino.Checked Then
            leyendacintura.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "obesidad morbida"
        End If
        Return leyendacintura.Text
    End Function






    Private Sub masculino_CheckedChanged(sender As Object, e As EventArgs) Handles masculino.CheckedChanged

    End Sub

    Private Sub sedentario_CheckedChanged(sender As Object, e As EventArgs) Handles sedentario.CheckedChanged

    End Sub

    Private Sub procesar_Click(sender As Object, e As EventArgs) Handles procesar.Click

        Try
            masacor.Text = Math.Round(ICM(), 2)
            leyenda.Text = Leyendamasa()
            indicintura.Text = Math.Round(ICE(), 2)
            leyendacintura.Text = Leyendaalcin()
            grasacorporal.Text = GC(ICM) & "%"
            porcentajegrasa.Text = Leyendagrasa()
            caloriadia.Text = CXD()

            descri.Text = nombre.Text & " " & apellido.Text & " con " & edad.Text & " años " & " debe consumir al dia " & caloriadia.Text & " calorias " & Chr(13) & Leyendaalcin()









        Catch ex As Exception
            MsgBox("Por favor rellene todos los campos")

        End Try

    End Sub

    Private Sub masacor_TextChanged(sender As Object, e As EventArgs) Handles masacor.TextChanged

    End Sub

    Private Sub indicintura_TextChanged(sender As Object, e As EventArgs) Handles indicintura.TextChanged

    End Sub

    Private Sub grasacorporal_TextChanged(sender As Object, e As EventArgs) Handles grasacorporal.TextChanged

    End Sub

    Private Sub caloriadia_TextChanged(sender As Object, e As EventArgs) Handles caloriadia.TextChanged

    End Sub

    Private Sub leyenda_Click(sender As Object, e As EventArgs) Handles leyenda.Click


    End Sub

    Private Sub edad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edad.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or edad.Text.Contains(".")) Then
            e.Handled = True
        End If

    End Sub

    Private Sub altura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles altura.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or altura.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub

    Private Sub peso_KeyPress(sender As Object, e As KeyPressEventArgs) Handles peso.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or peso.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub

    Private Sub cintura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cintura.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or cintura.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub

    Private Sub cadera_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cadera.KeyPress
        If Not (Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar)) _
            AndAlso (Not e.KeyChar = "." Or cadera.Text.Contains(".")) Then
            e.Handled = True
        End If
    End Sub

    Private Sub leyendacintura_Click(sender As Object, e As EventArgs) Handles leyendacintura.Click

    End Sub

    Private Sub porcentajegrasa_Click(sender As Object, e As EventArgs) Handles porcentajegrasa.Click

    End Sub


    Function Leyendagrasa()

        If (GC(ICM) < 8 And edad.Text < 18 And masculino.Checked) OrElse (GC(ICM) < 15 And edad.Text < 18 And femenino.Checked) Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "Delgadez"

        ElseIf GC(ICM) >= 8.1 And GC(ICM) <= 15.9 And edad.Text >= 18 And masculino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "peso normal"

        ElseIf GC(ICM) >= 15.1 And GC(ICM) <= 20.9 And edad.Text >= 18 And femenino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "peso normal"

        ElseIf GC(ICM) >= 16 And GC(ICM) <= 20.9 And edad.Text >= 18 And masculino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobrepeso ligero"

        ElseIf GC(ICM) >= 21 And GC(ICM) <= 25.9 And edad.Text >= 18 And femenino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobrepeso ligero"

        ElseIf GC(ICM) >= 21 And GC(ICM) <= 24.9 And edad.Text >= 18 And masculino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobrepeso"

        ElseIf GC(ICM) >= 26 And GC(ICM) <= 31.9 And edad.Text >= 18 And femenino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "sobrepeso"

        ElseIf GC(ICM) > 25 And edad.Text >= 18 And masculino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "Obesidad"

        ElseIf GC(ICM) > 32 And edad.Text >= 18 And femenino.Checked Then

            porcentajegrasa.Text = nombre.Text & " " & apellido.Text & " usted tiene " & "Obesidad"
        End If
        Return porcentajegrasa.Text
    End Function

    Private Sub limpiar_Click(sender As Object, e As EventArgs) Handles limpiar.Click
        nombre.Text = ""
        apellido.Text = ""
        edad.Text = ""
        altura.Text = ""
        peso.Text = ""
        cintura.Text = ""
        cadera.Text = ""
        masacor.Text = ""
        indicintura.Text = ""
        grasacorporal.Text = ""
        caloriadia.Text = ""

    End Sub

    Private Sub salir_Click(sender As Object, e As EventArgs) Handles salir.Click
        End
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles descri.Click

    End Sub
End Class
