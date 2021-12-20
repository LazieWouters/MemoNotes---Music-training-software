Imports System.Collections.Generic
Imports System.Text
Imports Microsoft.DirectX.DirectSound

''' <summary>
''' Capture device.
''' </summary>
Public Class SoundCaptureDevice
	Private m_id As Guid

	Private m_name As String

	Public ReadOnly Property IsDefault() As Boolean
		Get
			Return m_id = Guid.Empty
		End Get
	End Property

	''' <summary>
	''' Name of the device.
	''' </summary>
	Public ReadOnly Property Name() As String
		Get
			Return m_name
		End Get
	End Property

	Friend ReadOnly Property Id() As Guid
		Get
			Return m_id
		End Get
	End Property

	Friend Sub New(id As Guid, name As String)
		Me.m_id = id
		Me.m_name = name
	End Sub

	Public Shared Function GetDevices() As SoundCaptureDevice()
		Dim captureDevices As New CaptureDevicesCollection()
		Dim devices As New List(Of SoundCaptureDevice)()
		For Each captureDevice As DeviceInformation In captureDevices
			devices.Add(New SoundCaptureDevice(captureDevice.DriverGuid, captureDevice.Description))
		Next
		Return devices.ToArray()
	End Function

	Public Shared Function GetDefaultDevice() As SoundCaptureDevice
		Dim captureDevices As New CaptureDevicesCollection()
		Dim device As SoundCaptureDevice = Nothing
		For Each captureDevice As DeviceInformation In captureDevices
			If captureDevice.DriverGuid = Guid.Empty Then
				device = New SoundCaptureDevice(captureDevice.DriverGuid, captureDevice.Description)
				Exit For
			End If
		Next
		If device Is Nothing Then
			Throw New SoundCaptureException("Default capture device is not found")
		End If
		Return device
	End Function
End Class
