Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports SoundCapture

Public Partial Class SelectDeviceForm
	Inherits Form
	Private devices As SoundCaptureDevice()

	Public ReadOnly Property SelectedDevice() As SoundCaptureDevice
		Get
			Return devices(deviceNamesListBox.SelectedIndex)
		End Get
	End Property

	Public Sub New()
		InitializeComponent()
	End Sub

    Private Sub SelectDeviceForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            LoadDevices()
            If My.Settings.NovoValorDispositivoAudio <> 5000000 AndAlso Not OpcoesMemonotes.Visible Then
                deviceNamesListBox_DoubleClick2()
            Else
                Me.Opacity = 100
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub LoadDevices()

        Try

            deviceNamesListBox.Items.Clear()

            Dim defaultDeviceIndex As Integer = 0

            devices = SoundCaptureDevice.GetDevices()
            For i As Integer = 0 To devices.Length - 1
                deviceNamesListBox.Items.Add(devices(i).Name)
                If devices(i).IsDefault Then
                    defaultDeviceIndex = i
                End If
            Next

            If My.Settings.NovoValorDispositivoAudio <> 5000000 Then
                deviceNamesListBox.SelectedIndex = My.Settings.NovoValorDispositivoAudio
            Else
                deviceNamesListBox.SelectedIndex = defaultDeviceIndex
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub deviceNamesListBox_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles deviceNamesListBox.DoubleClick, okButton.Click

        deviceNamesListBox_DoubleClick2()

    End Sub

    Private Sub deviceNamesListBox_DoubleClick2()
        Try

            If deviceNamesListBox.SelectedIndex < 0 Then
                Return
            End If

            Me.DialogResult = DialogResult.OK
            My.Settings.NovoValorDispositivoAudio = deviceNamesListBox.SelectedIndex
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub
End Class
