Option Strict Off
Option Explicit On
Imports System.IO

Public Class MC ' MIDI Constant
    Public Const NOTE_OFF = &H80
    Public Const NOTE_ON = &H90
    Public Const POLY_KEY_PRESS = &HA0
    Public Const CONTROLLER_CHANGE = &HB0
    Public Const PROGRAM_CHANGE = &HC0
    Public Const CHANNEL_PRESSURE = &HD0
    Public Const PITCH_BEND = &HE0
    Public Const SYSEX = &HF0
    Public Const MTC_QFRAME = &HF1
    Public Const EOX = &HF7
    Public Const MIDI_CLOCK = &HF8
    Public Const MIDI_START = &HFA
    Public Const MIDI_CONTINUE = &HFB
    Public Const MIDI_STOP = &HFC
    Public Const MIDI_MAPPER = -1&
End Class


Public Class MidiEvent
    'variables locales de stockage des valeurs de propriétés
    Private m_D1H As Integer
    Private m_D1L As Integer
    Private m_D2 As Integer
    Private m_D3 As Integer
    Private m_Length As Integer
    Private m_Msg As Integer
    'Private m_Note_OFF As Long 
    'Private m_Note_Off_Pos As Long 

    Public Property D1H() As Integer
        Get
            Return m_D1H
        End Get
        Set(ByVal value As Integer)
            m_D1H = value
        End Set
    End Property

    Public Property D1L() As Integer
        Get
            Return m_D1L
        End Get
        Set(ByVal value As Integer)
            m_D1L = value
        End Set
    End Property

    Public Property D2() As Integer
        Get
            Return m_D2
        End Get
        Set(ByVal value As Integer)
            m_D2 = value
        End Set
    End Property

    Public Property D3() As Integer
        Get
            Return m_D3
        End Get
        Set(ByVal value As Integer)
            m_D3 = value
        End Set
    End Property

    Public Property Length() As Integer
        Get
            Return m_Length
        End Get
        Set(ByVal value As Integer)
            m_Length = value
        End Set
    End Property

    Public Property Msg() As Integer
        Get
            Return m_Msg
        End Get
        Set(ByVal value As Integer)
            m_Msg = value
        End Set
    End Property


End Class

Public Class MidiEvents
    Private mcol As Collection

    Public Function Add(ByVal D1H As Integer, ByVal D1L As Integer, _
                        ByVal D2 As Integer, ByVal D3 As Integer, _
                        ByVal Length As Integer) As MidiEvent

        Dim objNewMember As MidiEvent
        objNewMember = New MidiEvent

        objNewMember.D1H = D1H
        objNewMember.D1L = D1L
        objNewMember.D2 = D2
        objNewMember.D3 = D3
        objNewMember.Length = Length
        objNewMember.Msg = (D3 * &H10000) + _
                           (D2 * &H100) + _
                           (D1L - 1 + D1H)

        mcol.Add(objNewMember, "K" & D1H & "|" & D1L & "|" & D2)

        Add = objNewMember
        objNewMember = Nothing
    End Function

    Public Sub RemoveByKey(ByVal KeyEvent As String)
        mcol.Remove(KeyEvent)
    End Sub

    Public Sub RemoveByIndex(ByVal Index As Integer)
        mcol.Remove(Index)
    End Sub

    Public Function ContainKey(ByVal KeyEvent As String) As Boolean
        Return mcol.Contains(KeyEvent)
    End Function

    Public ReadOnly Property Count() As Long
        Get
            Count = mcol.Count
        End Get
    End Property

    Default Public ReadOnly Property KeyOrIndex(ByVal Index As Object) As MidiEvent
        Get
            Return mcol.Item(Index)
        End Get
    End Property

    Public Function GetEnumerator() As IEnumerator
        GetEnumerator = mcol.GetEnumerator
    End Function

    Public Sub New()
        MyBase.New()
        mcol = New Collection
    End Sub

    Protected Overrides Sub Finalize()
        mcol = Nothing
        MyBase.Finalize()
    End Sub

End Class

Public Class Tick
    Private m_MidiEvents As MidiEvents
    Private m_TickPos As Integer
    Private m_KeyToIndex As Integer

    Public Property KeyToIndex() As Integer
        Get
            Return m_KeyToIndex
        End Get
        Set(ByVal value As Integer)
            m_KeyToIndex = value
        End Set
    End Property


    Public Property TickPos() As Integer
        Get
            Return m_TickPos
        End Get
        Set(ByVal value As Integer)
            m_TickPos = value
        End Set
    End Property

    Public Property MidiEvents() As MidiEvents
        Get
            If m_MidiEvents Is Nothing Then
                m_MidiEvents = New MidiEvents
            End If
            MidiEvents = m_MidiEvents
        End Get
        Set(ByVal value As MidiEvents)
            m_MidiEvents = value
        End Set
    End Property

    Protected Overrides Sub Finalize()
        m_MidiEvents = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub New()
        MyBase.New()
        m_MidiEvents = New MidiEvents
    End Sub
End Class

Public Class Ticks
    Private mCol As Collection

    Public Function Add(ByVal TickPos As Integer) As Tick
        Dim objNewTick As Tick
        objNewTick = New Tick

        If mCol.Count Then
            If Not HasKey("K" & TickPos) Then
                If (mCol(mCol.Count).TickPos < TickPos) Then
                    ' Le New Tick a une Pos Supérieure au dernier Tick de mCol
                    'Add = New Tick
                    objNewTick.TickPos = TickPos
                    objNewTick.KeyToIndex = mCol.Count + 1
                    mCol.Add(objNewTick, "K" & TickPos)
                    'Add = objNewTick
                    'objNewTick = Nothing
                Else
                    Dim i As Long
                    For i = mCol.Count To 1 Step -1
                        ' recaler les index ici
                        mCol(i).KeyToIndex = i + 1
                        If (mCol(i).TickPos < TickPos) Then
                            'Add = New Tick
                            objNewTick.TickPos = TickPos
                            objNewTick.KeyToIndex = i + 1
                            mCol.Add(objNewTick, "K" & TickPos, , i)
                            ' comme on insère after i on recale i
                            mCol(i).KeyToIndex = i
                            Add = objNewTick
                            objNewTick = Nothing
                            Exit Function
                        End If
                    Next
                    ' C'est le premier Tick de mcol
                    'Add = New Tick
                    objNewTick.TickPos = TickPos
                    objNewTick.KeyToIndex = 1
                    mCol.Add(objNewTick, "K" & TickPos, 1)
                End If
            Else
                Add = mCol("K" & TickPos)
                objNewTick = Nothing
                Exit Function
            End If
        Else
            objNewTick.TickPos = TickPos
            objNewTick.KeyToIndex = 1
            mCol.Add(objNewTick, "K" & TickPos)
        End If

        Add = objNewTick
        objNewTick = Nothing

    End Function

    Public Sub RemoveByKey(ByVal Key As String)
        mCol.Remove(Key)
        Dim Item As Tick
        Dim i As Integer : i = 1
        ' attention la collection commence à 1 (base1)
        For Each Item In mCol
            Item.KeyToIndex = i
            i = i + 1
        Next Item
    End Sub

    Public Sub RemoveByIndex(ByVal Index As Integer)
        mCol.Remove(Index)
        Dim Item As Tick
        Dim i As Integer : i = 1
        ' attention la collection commence à 1 (base 1)
        For Each Item In mCol
            Item.KeyToIndex = i
            i = i + 1
        Next Item
    End Sub

    Default Public ReadOnly Property Key(ByVal KPos As String) As Tick
        Get
            Return mCol.Item(KPos)
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Count = mCol.Count
        End Get
    End Property

    Public Function GetEnumerator() As IEnumerator
        GetEnumerator = mCol.GetEnumerator
    End Function

    Public Function HasKey(ByVal Key As String) As Boolean
        Return mCol.Contains(Key)
    End Function

    Public Sub New()
        mCol = New Collection
    End Sub

    Protected Overrides Sub Finalize()
        mCol = Nothing
        MyBase.Finalize()
    End Sub
End Class

Public Class Track
    Private m_Ticks As Ticks
    Private m_TrackName As String
    Private m_Mute As Boolean
    Private m_Canal As Integer


    Public Sub AddMidiEvent(ByRef TickPos As Integer, _
                            ByRef D1H As Integer, ByRef D1L As Integer, _
                            ByRef D2 As Integer, ByRef D3 As Integer, _
                            ByRef Lenght As Integer)

        Dim vEvent As MidiEvent
        Dim vTick As Tick

        vTick = m_Ticks.Add(TickPos)

        If vTick.MidiEvents.ContainKey("K" & D1H & "|" & D1L & "|" & D2) Then
            'NOP
        Else
            vEvent = vTick.MidiEvents.Add(D1H, D1L, D2, D3, Lenght)
        End If

    End Sub

    Public Sub RemMidiEvent(ByRef TickPos As Integer, _
                            ByRef D1H As Integer, ByRef D1L As Integer, _
                            ByRef D2 As Integer)

        Dim Evs As MidiEvents
        Dim KeyPos As String : KeyPos = "K" & TickPos
        Dim KeyEvent As String : KeyEvent = "K" & D1H & "|" & D1L & "|" & D2

        Evs = m_Ticks(KeyPos).MidiEvents
        Debug.Print("Key event = " & KeyEvent)
        If Evs.ContainKey(KeyEvent) Then Evs.RemoveByKey(KeyEvent)

        ' Reste t il des MIDI Event ? si non on supprime le tick
        If Evs.Count = 0 Then
            m_Ticks.RemoveByKey(KeyPos)
        End If

    End Sub

    Public Property Ticks() As Ticks
        Get
            If m_Ticks Is Nothing Then
                m_Ticks = New Ticks
            End If
            Return m_Ticks
        End Get
        Set(ByVal Value As Ticks)
            m_Ticks = Value
        End Set
    End Property

    Public Property Mute() As Boolean
        Get
            Return m_Mute
        End Get
        Set(ByVal Value As Boolean)
            m_Mute = Value
        End Set
    End Property

    Public Property Canal() As Integer
        Get
            Return m_Canal
        End Get
        Set(ByVal Value As Integer)
            m_Canal = Value
        End Set
    End Property

    Public Property TrackName() As String
        Get
            Return m_TrackName
        End Get
        Set(ByVal Value As String)
            m_TrackName = Value
        End Set
    End Property

    Public Sub Clear()

        m_Ticks = Nothing
        m_Ticks = New Ticks

    End Sub

    Public Sub New()
        MyBase.New()
        m_Ticks = New Ticks
    End Sub

    Protected Overrides Sub Finalize()
        m_Ticks = Nothing
        MyBase.Finalize()
    End Sub
End Class

Public Class Tracks
    Implements System.Collections.IEnumerable

    Private mCol As Collection
    Private m_Tempo As Double

    Private Structure NoteLenght
        Dim ON_OFF As Byte
        Dim TickPosON As Integer
        Dim TickPosOFF As Integer
        Dim Lenght As Integer
        Dim Index As Integer
    End Structure

    Private MFBuffer() As Byte                      ' Buffer va recevoir le MIDI file
    Private MFFilePos As Integer                    ' Position du MidiFile dans le MFBuffer
    Private CurTrack As Integer                     ' Piste en cours du SMF
    Private Format As Integer                       ' Type du Midifile 0 - 1 - 2
    Private NoteLenghtTable(31, 127) As NoteLenght  ' pour le SMF

    'Private DBUG As Boolean

    Private running_status(255) As Byte
    Private last_running_status As Byte
    Private trnr As Integer

    Public Sub TrackAdd(ByVal Name As String)
        Dim ObjNewTrack As Track
        ObjNewTrack = New Track

        ObjNewTrack.TrackName = Name

        mCol.Add(ObjNewTrack, Name)

        ObjNewTrack = Nothing

    End Sub

    Public Sub RemoveTrackByName(ByVal TrackName As String)
        mCol.Remove(TrackName)
    End Sub

    Public ReadOnly Property MF_Tempo() As Double
        Get
            'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
            'Syntax: Debug.Print X.MF_Tempo
            Return m_Tempo
        End Get
    End Property

    Default Public ReadOnly Property Key(ByVal TrackName As String) As Track
        Get
            Return mCol.Item(TrackName)
        End Get
    End Property


    Public Sub New()
        MyBase.New()
        mCol = New Collection
    End Sub

    Protected Overrides Sub Finalize()
        mCol = Nothing
        MyBase.Finalize()
    End Sub

    Public Function GetEnumerator() As System.Collections.IEnumerator _
                                    Implements System.Collections.IEnumerable.GetEnumerator
        Return mCol.GetEnumerator
    End Function

    Dim fs As FileStream
    Dim br As BinaryReader

    Public Sub OpenMidiFile(ByVal strFileName As String)
        'Dim FileHeaderId As Byte() = New System.Text.ASCIIEncoding().GetBytes("MThd")
        'Dim TrkHeaderId As Byte() = New System.Text.ASCIIEncoding().GetBytes("MTrk")
        Dim MidiFilelen As Integer ' longueur du SMF
        Dim Trks As Integer
        Dim Kres As Integer ' Coef pour résolution
        Dim running_status As Byte
        Dim Tempotrack As Integer
        Dim CurTrackTickPos As Integer
        Dim dw As Integer ' premier delta
        Dim iEv As Integer ' Index pour les events
        Dim MidiByte As Byte
        Dim Data1, Data2 As Byte

        Dim TickLen As Integer ' Longueur en Tick
        Dim Delta As Integer

        Dim Trk As Track
        Dim Itm As MidiEvent

        ' Efface le contenu des pistes
        For Each Trk In mCol
            If Trk.Ticks.Count Then Trk.Clear()
        Next Trk

        MFFilePos = 0

        fs = New FileStream(strFileName, FileMode.Open, FileAccess.Read)

        br = New BinaryReader(fs)
        MidiFilelen = fs.Length
        MFBuffer = br.ReadBytes(MidiFilelen)
        br = Nothing
        fs.Close()

        ' Dans l'ordre on lit le midifile
        If System.Text.Encoding.Default.GetString(MFBuffer, 0, 4) = "MThd" Then
            'MessageBox.Show("Ok")
        Else
            MessageBox.Show("Error File Header MThd")
            Exit Sub
        End If

        ' on avance br de 4 correspond à la longueur du chunk sur 4 octets
        ' celle du Header chunck étant de 6 -> 00H 00H 00H 06H ok ?

        'Get File Format - Extrait le Format - 2 octets
        Dim Temp() As Byte = {MFBuffer(9), MFBuffer(8)}
        Format = BitConverter.ToInt16(Temp, 0)

        'Get Number of tracks - Extrait le nombre de piste - 2 octets
        Temp = New Byte() {MFBuffer(11), MFBuffer(10)}
        Trks = BitConverter.ToInt16(Temp, 0)

        'Get Division ou resolution division négative si Frame Based
        'Extrait la division - 2 Octets et transforme ds la resolution actuelle
        Temp = New Byte() {MFBuffer(13), MFBuffer(12)}
        Kres = BitConverter.ToInt16(Temp, 0) / &H30
        'Debug.Print("Kres = " & Kres)

        running_status = &HFES

        MFFilePos = 14

        ' On initialise le running_status
        running_status = &HFES

        Dim HeaderOk As Boolean

        ' On va parcourir le MidiFile
        Do While MFFilePos < MidiFilelen '
            ' vérifie MTrk
            If MFFilePos >= MidiFilelen Then
                Exit Do
            End If
            'check le Header Track
            If MFFilePos < MidiFilelen - 4 Then
                If System.Text.Encoding.Default.GetString(MFBuffer, MFFilePos, 4) = "MTrk" Then
                    HeaderOk = True
                Else
                    HeaderOk = False
                End If
            End If

            If HeaderOk Then
                If Format = 1 Then
                    MFFilePos += 4
                    MFFilePos += 4
                    dw = GetVarLen()
                    'determine la piste en cours - la première étant celle du tempo
                    If CurTrack = 0 And Tempotrack = False Then
                        Tempotrack = True
                    Else
                        CurTrack += 1
                        Trk = mCol.Item(CurTrack)
                        CurTrackTickPos = 0 + dw
                    End If
                ElseIf Format = 0 Then
                    MFFilePos += 4
                    MFFilePos += 4
                    dw = GetVarLen()
                    CurTrack += 1
                    Trk = mCol.Item(CurTrack)
                    CurTrackTickPos = 0 + dw
                End If
                HeaderOk = False
            End If

            ' Récupère Un Octet
            MidiByte = MFBuffer(MFFilePos) : MFFilePos += 1
            Select Case (MidiByte) ' 3 types d'évenements
                Case &HFFS
                    MetaEvent()
                    running_status = &HFES ' &annule le running status
                Case &HF0S
                    SysexEvent()
                Case &HF7S
                    'MsgBox "SysEx"
                Case &HF1S
                Case &HF2S
                Case &HF3S
                Case &HF4S
                Case &HF5S
                Case &HF6S
                    'MsgBox "system Common"
                Case &HF7S
                Case &HF8S
                Case &HF9S
                Case &HFAS
                Case &HFBS
                Case &HFCS
                Case &HFDS
                Case &HFES
                    'MsgBox "system real time"
                Case Else 'midi events
                    Select Case (MidiByte And &HF0S)
                        Case MC.NOTE_OFF '&H80 - 128
                            Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                            Data2 = MFBuffer(MFFilePos) : MFFilePos += 1
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MidiByte), "00") & " " & VB6.Format(Hex(Data1), "00") & " " & VB6.Format(Hex(Data2), "00") & " ", TAB, MFFilePos - 3, TAB, MFFilePos, TAB, "Note OFF Ch=" & (MidiByte - MC.NOTE_OFF + 1) & " N=" & Data1 & " V=" & Data2))
                            If Format = 0 Then CurTrack = (MidiByte - MC.NOTE_OFF) + 1 : Trk = mCol.Item(CurTrack)
                            NoteLenghtTable(CurTrack, Data1).Lenght = (CurTrackTickPos \ Kres) - NoteLenghtTable(CurTrack, Data1).TickPosON
                            iEv = NoteLenghtTable(CurTrack, Data1).Index
                            Itm = Trk.Ticks("K" & NoteLenghtTable(CurTrack, Data1).TickPosON).MidiEvents(iEv)
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(TAB, TAB, TAB, "Index = " & iEv & " PosOn = " & NoteLenghtTable(CurTrack, Data1).TickPosON))
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(TAB, TAB, TAB, "Track = " & CurTrack & " Lenght = " & (CurTrackTickPos \ Kres) - NoteLenghtTable(CurTrack, Data1).TickPosON))
                            Itm.Length = NoteLenghtTable(CurTrack, Data1).Lenght
                            If Itm.Length = 0 Then Itm.Length = 1 ' mini lenght = 1
                        Case MC.NOTE_ON '&H90 - 144
                            Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                            Data2 = MFBuffer(MFFilePos) : MFFilePos += 1
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MidiByte), "00") & " " & VB6.Format(Hex(Data1), "00") & " " & VB6.Format(Hex(Data2), "00") & " ", TAB, MFFilePos - 3, TAB, MFFilePos, TAB, "Note ON Ch=" & (MidiByte - MC.NOTE_ON + 1) & " N=" & Data1 & " V=" & Data2))
                            If Format = 0 Then CurTrack = (MidiByte - MC.NOTE_ON) + 1 : Trk = mCol.Item(CurTrack)
                            If Data2 = 0 Then
                                ' note_off -> on ajuste la longueueur de la note
                                NoteLenghtTable(CurTrack, Data1).Lenght = (CurTrackTickPos \ Kres) - NoteLenghtTable(CurTrack, Data1).TickPosON
                                iEv = NoteLenghtTable(CurTrack, Data1).Index
                                Itm = Trk.Ticks("K" & NoteLenghtTable(CurTrack, Data1).TickPosON).MidiEvents(iEv)
                                Itm.Length = NoteLenghtTable(CurTrack, Data1).Lenght
                                If Itm.Length = 0 Then Itm.Length = 1
                            Else
                                ' Note_on -> on crée une note
                                Trk.AddMidiEvent(CurTrackTickPos \ Kres, MC.NOTE_ON, MidiByte - MC.NOTE_ON + 1, CInt(Data1), CInt(Data2), 1)
                                NoteLenghtTable(CurTrack, Data1).ON_OFF = True
                                NoteLenghtTable(CurTrack, Data1).TickPosON = CurTrackTickPos \ Kres
                                NoteLenghtTable(CurTrack, Data1).Index = Trk.Ticks("K" & (CurTrackTickPos \ Kres)).MidiEvents.Count
                            End If
                        Case MC.POLY_KEY_PRESS ' &HA0 - 160
                            Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                            Data2 = MFBuffer(MFFilePos) : MFFilePos += 1
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(Hex(MidiByte) & " " & Hex(Data1) & " " & Hex(Data2), TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Key Press Ch=" & (MidiByte - MC.POLY_KEY_PRESS + 1) & " N=" & Data1 & " V=" & Data2))
                            If Format = 0 Then CurTrack = (MidiByte - MC.POLY_KEY_PRESS) + 1 : Trk = mCol.Item(CurTrack)
                            Trk.AddMidiEvent(CurTrackTickPos \ Kres, MC.POLY_KEY_PRESS, MidiByte - MC.POLY_KEY_PRESS + 1, CInt(Data1), CInt(Data2), 0)
                        Case MC.CONTROLLER_CHANGE ' &HB0 - 176
                            Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                            Data2 = MFBuffer(MFFilePos) : MFFilePos += 1
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(Hex(MidiByte) & " " & Hex(Data1) & " " & Hex(Data2), TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Ctrl Change Ch=" & (MidiByte - MC.CONTROLLER_CHANGE + 1) & " Ctrl=" & Data1 & " Val=" & Data2))
                            If Format = 0 Then CurTrack = (MidiByte - MC.CONTROLLER_CHANGE) + 1 : Trk = mCol.Item(CurTrack)
                            Trk.AddMidiEvent(CurTrackTickPos \ Kres, MC.CONTROLLER_CHANGE, MidiByte - MC.CONTROLLER_CHANGE + 1, CInt(Data1), CInt(Data2), 0)
                            Select Case Data1
                                Case 7
                                    '                      mCol(CurTrack).Volume = Data2
                                Case 10
                                Case Else
                            End Select
                        Case MC.PROGRAM_CHANGE
                            Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MFBuffer(MFFilePos - 2)), "00") & " " & VB6.Format(Hex(MFBuffer(MFFilePos - 1)), "00"), TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Prg Change Canal Ch=" & (MidiByte - MC.PROGRAM_CHANGE + 1) & " PG=" & Data1))
                            If Format = 0 Then CurTrack = (MidiByte - MC.PROGRAM_CHANGE) + 1 : Trk = mCol.Item(CurTrack)
                            Trk.AddMidiEvent(CurTrackTickPos \ Kres, MC.PROGRAM_CHANGE, MidiByte - MC.PROGRAM_CHANGE + 1, CInt(Data1), 0, 0)
                            '               mCol(CurTrack).Program = Data1 + 1
                        Case MC.CHANNEL_PRESSURE
                            Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(Hex(MidiByte) & " " & Hex(Data1), TAB, MFFilePos, TAB, "Channel pressure Ch=" & (MidiByte - MC.CHANNEL_PRESSURE + 1) & " Val=" & Data1))
                            If Format = 0 Then CurTrack = (MidiByte - MC.CHANNEL_PRESSURE) + 1 : Trk = mCol.Item(CurTrack)
                            Trk.AddMidiEvent(CurTrackTickPos \ Kres, MC.CHANNEL_PRESSURE, MidiByte - MC.CHANNEL_PRESSURE + 1, CInt(Data1), 0, 0)
                        Case MC.PITCH_BEND ' seulement 7 bits utiles par octet
                            Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                            Data2 = MFBuffer(MFFilePos) : MFFilePos += 1
                            'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(Hex(MidiByte) & " " & Hex(Data1) & " " & Hex(Data2), TAB, MFFilePos - 3, TAB, MFFilePos, TAB, "Pitch Bend Ch=" & (MidiByte - MC.PITCH_BEND + 1) & " Lsb=" & Data1 & " Msb=" & Data2))
                            If Format = 0 Then CurTrack = (MidiByte - MC.PITCH_BEND) + 1 : Trk = mCol.Item(CurTrack)
                            Trk.AddMidiEvent(CurTrackTickPos \ Kres, MC.PITCH_BEND, MidiByte - MC.PITCH_BEND + 1, CInt(Data1), CInt(Data2), 0)
                        Case Else
                            If running_status <> last_running_status Then
                                last_running_status = running_status
                            End If
                            If (MidiByte And &H80) = 0 Then ' running
                                Select Case (running_status And &HF0)
                                    Case MC.NOTE_OFF
                                        Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                                        'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MidiByte), "00") & " " & VB6.Format(Hex(Data1), "00") & " ", TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Note Off Ch=" & (CShort(running_status - CShort(running_status And &HF0S)) + 1) & " N=" & MidiByte & " V=" & Data1 & " (Running (" & Hex(running_status) & "))"))
                                        If Format = 0 Then CurTrack = CShort(running_status - CShort(running_status And &HF0S)) + 1 : Trk = mCol.Item(CurTrack)
                                        NoteLenghtTable(CurTrack, MidiByte).Lenght = (CurTrackTickPos \ Kres) - NoteLenghtTable(CurTrack, MidiByte).TickPosON
                                        iEv = NoteLenghtTable(CurTrack, MidiByte).Index
                                        Itm = Trk.Ticks("K" & NoteLenghtTable(CurTrack, MidiByte).TickPosON).MidiEvents(iEv)
                                        'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(TAB, TAB, TAB, "Index = " & iEv & " PosOn = " & NoteLenghtTable(CurTrack, MidiByte).TickPosON))
                                        'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(TAB, TAB, TAB, "Track = " & CurTrack & " Lenght = " & (CurTrackTickPos \ Kres) - NoteLenghtTable(CurTrack, MidiByte).TickPosON))
                                        Itm.Length = NoteLenghtTable(CurTrack, MidiByte).Lenght
                                        If Itm.Length = 0 Then Itm.Length = 1
                                    Case MC.NOTE_ON
                                        Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                                        'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MidiByte), "00") & " " & VB6.Format(Hex(Data1), "00") & " ", TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Note On Ch=" & (CShort(running_status - CShort(running_status And &HF0S)) + 1) & " N=" & MidiByte & " V=" & Data1 & " (Running (" & Hex(running_status) & "))"))
                                        If Format = 0 Then CurTrack = CShort(running_status - CShort(running_status And &HF0S)) + 1 : Trk = mCol.Item(CurTrack)
                                        If Data1 = 0 Then
                                            ' note_off -> on ajuste la longueueur de la note
                                            NoteLenghtTable(CurTrack, MidiByte).Lenght = (CurTrackTickPos \ Kres) - NoteLenghtTable(CurTrack, MidiByte).TickPosON
                                            iEv = NoteLenghtTable(CurTrack, MidiByte).Index
                                            Itm = Trk.Ticks("K" & NoteLenghtTable(CurTrack, MidiByte).TickPosON).MidiEvents(iEv)
                                            Itm.Length = NoteLenghtTable(CurTrack, MidiByte).Lenght
                                            If Itm.Length = 0 Then Itm.Length = 1
                                        Else
                                            ' Note_on -> on crée une note
                                            Trk.AddMidiEvent(CurTrackTickPos \ Kres, MC.NOTE_ON, CShort(running_status - CShort(running_status And &HF0S)) + 1, CInt(MidiByte), CInt(Data1), 1)
                                            NoteLenghtTable(CurTrack, MidiByte).ON_OFF = True
                                            NoteLenghtTable(CurTrack, MidiByte).Index = Trk.Ticks("K" & (CurTrackTickPos \ Kres)).MidiEvents.Count
                                            NoteLenghtTable(CurTrack, MidiByte).TickPosON = CurTrackTickPos \ Kres
                                        End If
                                        '                     Case &HA0, &HD0, &HE0
                                        '                        Data1 = MFBuffer(MFFilePos): MFFilePos += 1
                                    Case MC.CONTROLLER_CHANGE '&HB0 ' Control Change
                                        Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                                        'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MidiByte), "00") & " " & VB6.Format(Hex(Data1), "00") & " ", TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Ctrl Change Ch=" & (CShort(running_status - CShort(running_status And &HF0S)) + 1) & " Ctrl=" & MidiByte & " Val=" & Data1 & " (Running (" & Hex(running_status) & "))"))
                                        If Format = 0 Then CurTrack = CShort(running_status - CShort(running_status And &HF0S)) + 1 : Trk = mCol.Item(CurTrack)
                                        Trk.AddMidiEvent(CurTrackTickPos \ Kres, &HB0S, CShort(running_status - CShort(running_status And &HF0S)) + 1, CInt(MidiByte), CInt(Data1), 0)

                                    Case Else
                                        Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                                        'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MidiByte), "00") & " " & VB6.Format(Hex(Data1), "00") & " ", TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Ctrl Else Ch=" & (CShort(running_status - CShort(running_status And &HF0S)) + 1) & " Ctrl=" & MidiByte & " Val=" & Data1 & " (Running (" & Hex(running_status) & "))"))
                                        If Format = 0 Then CurTrack = CShort(running_status - CShort(running_status And &HF0S)) + 1 : Trk = mCol.Item(CurTrack)
                                        Trk.AddMidiEvent(CurTrackTickPos \ Kres, running_status And &HF0S, CShort(running_status - CShort(running_status And &HF0S)) + 1, CInt(MidiByte), CInt(Data1), 0)
                                End Select
                                Delta = GetVarLen()
                                CurTrackTickPos = CurTrackTickPos + Delta
                                If TickLen < CurTrackTickPos Then TickLen = CurTrackTickPos
                            End If
                    End Select
                    If (MidiByte And &H80) <> 0 Then ' Not running
                        Delta = GetVarLen()
                        '               Debug.Print "Delta = " & delta & " MFFilepos = " & MFFilePos
                        CurTrackTickPos = CurTrackTickPos + Delta
                        '               Debug.Print "CurTrackTickPos = " & CurTrackTickPos
                        running_status = MidiByte
                        last_running_status = MidiByte
                        If TickLen < CurTrackTickPos Then TickLen = CurTrackTickPos
                    End If
            End Select
        Loop


        'mvarMF_Lenght = (((TickLen \ Kres) \ 192) * 192) + 192
        'mvarMF_TrackCount = CurTrack
        'mvarMF_Type = Format
        '
        'Call ResetBtn_Click
        'MsgBox(" MF chargé ")

    End Sub

    '***************************************************************************************
    ' GetVarLen
    ' Lit un nombre d'Octet variable (max 4 octets) et renvoi la valeur
    ' Pour tester s'il y a un autre octet à lire on vérifie le bit 7
    '***************************************************************************************
    Private Function GetVarLen() As Integer
        Dim Value As Integer
        If MFBuffer(MFFilePos) <= &H7F Then ' 1 seul octet
            Value = MFBuffer(MFFilePos)
            MFFilePos += 1
            '   If DBUG Then Debug.Print Format(Hex$(MFBuffer(P - 1)), "00") & " "; Tab; P - 1; Tab; P; Tab; "Variable Len " & Value & " (1 octet)"
            Return Value
            Exit Function
        ElseIf MFBuffer(MFFilePos + 1) <= &H7F Then  ' 2 octets
            Value = ((MFBuffer(MFFilePos) And &H7F) * &H80) Or (MFBuffer(MFFilePos + 1) And &H7F)
            MFFilePos += 2
            '   If DBUG Then Debug.Print Format(Hex$(MFBuffer(P - 2)), "00") & " " & Format(Hex$(MFBuffer(P - 1)), "00") & " "; Tab; P - 2; Tab; P; Tab; "Variable Len " & Value & " (2 octets)"
            Return Value
            Exit Function
        ElseIf MFBuffer(MFFilePos + 2) <= &H7F Then  ' 3 octets
            Value = (MFBuffer(MFFilePos) And &H7F) * &H80
            Value = (Value Or (MFBuffer(MFFilePos + 1) And &H7F)) * &H80
            Value = Value Or (MFBuffer(MFFilePos + 2) And &H7F)
            MFFilePos += 3
            '   If DBUG Then Debug.Print Format(Hex$(MFBuffer(P - 3)), "00") & " " & Format(Hex$(MFBuffer(P - 2)), "00") & " " & Format(Hex$(MFBuffer(P - 1)), "00") & " "; Tab; P - 3; Tab; P; Tab; "Variable Len " & Value & " (3 octets)"
            Return Value
            Exit Function
        ElseIf MFBuffer(MFFilePos + 3) <= &H7F Then  ' 4 octets
            Value = (MFBuffer(MFFilePos) And &H7F) * &H80
            Value = (Value Or (MFBuffer(MFFilePos + 1) And &H7F)) * &H80
            Value = (Value Or (MFBuffer(MFFilePos + 2) And &H7F)) * &H80
            Value = (Value Or (MFBuffer(MFFilePos + 3) And &H7F))
            MFFilePos += 4
            '   If DBUG Then Debug.Print Format(Hex$(MFBuffer(P - 4)), "00") & " " & Format(Hex$(MFBuffer(P - 3)), "00") & " " & Format(Hex$(MFBuffer(P - 2)), "00") & " " & Format(Hex$(MFBuffer(P - 1)), "00") & " "; Tab; P - 4; Tab; P; Tab; "Variable Len " & Value & " (4 octets)"
            Return Value
            Exit Function
        End If
    End Function

    '***************************************************************************************
    ' MetaEvent
    '***************************************************************************************
    Sub MetaEvent()
        Dim c As Byte
        Dim X As Integer
        Dim l As Integer
        Dim Data1 As Byte
        Dim Data2 As Byte

        'UPGRADE_NOTE: Stra été mis à niveau vers Str_Renamed. Cliquez ici : 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Str_Renamed As String : Str_Renamed = ""

        'Call Get8(c, MFFilePos, LLeft)
        c = MFBuffer(MFFilePos) : MFFilePos += 1
        'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout("FF " & Hex(c) & " ", TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Meta Event Track = " & trnr))

        'Dim Temp As Integer
        Dim Tbl(3) As Byte 'Tableau temporaire pour copie de  4 octet ' variable temporaire pour le tempo
        Select Case c
            Case &H0S
            Case &H1S To &H2S ' text event
                l = GetVarLen()
                For X = 0 To (l - 1)
                    c = MFBuffer(MFFilePos) : MFFilePos += 1
                    'Str = Str & Chr(MFBuffer(MFFilePos - 1))
                    'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(c), "00"), TAB, MFFilePos - 1, TAB, MFFilePos, TAB, "Meta Event Track = " & trnr & " ...HEXA = " & Chr(MFBuffer(MFFilePos - 1))))
                Next
                l = GetVarLen()
                '      track_delta(trnr) = track_delta(trnr) + l
                'Case &H2 ' copyright notice
            Case &H3S To &H8S ' sequence / track name
                l = GetVarLen()
                For X = 0 To (l - 1)
                    '            Call Get8(c, MFFilePos, LLeft)
                    c = MFBuffer(MFFilePos) : MFFilePos += 1
                    Str_Renamed = Str_Renamed & Chr(MFBuffer(MFFilePos - 1))
                    'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(c), "00"), TAB, MFFilePos - 1, TAB, MFFilePos, TAB, "Meta Event Track = " & trnr & " ...Chr = " & Chr(MFBuffer(MFFilePos - 1))))
                Next
                If Format = 0 Then
                    '         SeqNameLbl.Caption = Str
                Else
                    If CurTrack = 0 Then
                        '            SeqNameLbl.Caption = Str
                    Else
                        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet mCol().TrackName. Cliquez ici : 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mCol.Item(CurTrack).TrackName = Str_Renamed
                    End If
                    '         Debug.Print CurTrack
                End If
                '      Debug.Print CStr(Str)
                l = GetVarLen()
                '      track_delta(trnr) = track_delta(trnr) + l
                'Case &H4 ' Instrument name
                'Case &H5 ' Chant - Lyrics
                'Case &H6 ' Cue Point
                'Case &H7 ' Bruitage ->>>
                'Case &H8
            Case &H9S
            Case &HAS
            Case &HBS

            Case &HCS
            Case &HDS
            Case &HES
            Case &H21S
                l = GetVarLen()
                For X = 0 To (l - 1)
                    c = MFBuffer(MFFilePos) : MFFilePos += 1
                    Str_Renamed = Str_Renamed & Chr(MFBuffer(MFFilePos - 1))
                    'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(c), "00"), TAB, MFFilePos - 1, TAB, MFFilePos, TAB, "Meta Event Track = " & trnr & " ...Chr = " & Chr(MFBuffer(MFFilePos - 1))))
                Next
                l = GetVarLen()
                '      track_delta(trnr) = track_delta(trnr) + l
            Case &HFS
                'If DBUG Then Debug.Print "FF"; Hex$(c); Tab; MFFilePos; Tab; "Variable Len " & c
            Case &H2FS ' Fin de Piste - End of Track
                c = MFBuffer(MFFilePos) : MFFilePos += 1
                '      Debug.Print Format(Hex$(c), "00") & " "; Tab; MFFilePos - 1; Tab; MFFilePos; Tab; "Fin de Piste - FF 2F 00"
                'Surtout pas de Delta time apres le Track End
            Case &H51S ' tempo
                ' Lit la longueur du bloc - 3 Octets
                '      c = MFBuffer(MFFilePos): MFFilePos += 1
                MFFilePos += 1
                'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MFBuffer(MFFilePos - 1)), "00") & " ", TAB, MFFilePos - 1, TAB, MFFilePos, TAB, "Tempo - FF 51 03"))
                ' Lit les 3 Octets du Tempo en µs par noire
                Tbl(0) = MFBuffer(MFFilePos + 2) 'Zone poid faible
                Tbl(1) = MFBuffer(MFFilePos + 1)
                Tbl(2) = MFBuffer(MFFilePos)
                Tbl(3) = 0 'Zone de poid fort
                MFFilePos = MFFilePos + 3
                Dim Temp() As Byte = {Tbl(0), Tbl(1), Tbl(2), Tbl(3)}
                m_Tempo = (60 * 1000000) / BitConverter.ToInt32(Temp, 0)

                'RtlMoveMemory(Temp, Tbl(0), 4)
                'm_Tempo = (60 * 1000000) / Temp
                '      If DBUG Then Debug.Print Format(Hex$(MFBuffer(MFFilePos - 3)), "00") & " " & Format(Hex$(MFBuffer(MFFilePos - 2)), "00") & " "; Format(Hex$(MFBuffer(MFFilePos - 1)), "00") & " "; Tab; MFFilePos - 3; Tab; MFFilePos; Tab; "Tempo = " & Tempo1.Tempo
                'Delta time juste apres
                l = GetVarLen()
            Case &H54S
                l = GetVarLen()
                For X = 0 To (l - 1)
                    c = MFBuffer(MFFilePos) : MFFilePos += 1
                    Str_Renamed = Str_Renamed & Chr(MFBuffer(MFFilePos - 1))
                    'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(c), "00"), TAB, MFFilePos - 1, TAB, MFFilePos, TAB, "Meta Event 54 Track = " & trnr & " ...Byte = " & Chr(MFBuffer(MFFilePos - 1))))
                Next
                l = GetVarLen()
            Case &H58S ' time signature
                l = GetVarLen()
                c = MFBuffer(MFFilePos) : MFFilePos += 1
                'MainForm.LblNum = Numerateur
                c = MFBuffer(MFFilePos) : MFFilePos += 1
                '        MainForm.Tempo1.Denum = (c ^ 2)
                'MainForm.LblDenom = Denominateur
                c = MFBuffer(MFFilePos) : MFFilePos += 1
                c = MFBuffer(MFFilePos) : MFFilePos += 1
                'Debug.Print "Track " & trnr; Tab; Numerateur; Tab; Denominateur
                'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MFBuffer(MFFilePos - 3)), "00") & " " & VB6.Format(Hex(MFBuffer(MFFilePos - 2)), "00") & " " & VB6.Format(Hex(MFBuffer(MFFilePos - 1)), "00") & " ", TAB, MFFilePos - 3, TAB, MFFilePos, TAB, "Time signature"))
                'Call Get32(l, MFFilePos, LLeft)
                'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(MFBuffer(MFFilePos - 4)), "00") & " " & VB6.Format(Hex(MFBuffer(MFFilePos - 3)), "00") & " " & VB6.Format(Hex(MFBuffer(MFFilePos - 2)), "00") & " " & VB6.Format(Hex(MFBuffer(MFFilePos - 1)), "00") & " ", TAB, MFFilePos - 4, TAB, MFFilePos, TAB, "Mesure + metronome = " & l))
                'Delta time juste apres
                l = GetVarLen()
            Case &H59S 'Key Signature
                c = MFBuffer(MFFilePos) : MFFilePos += 1
                Data1 = MFBuffer(MFFilePos) : MFFilePos += 1
                If CurTrack = 0 Then
                    'tempos Track ?
                Else
                    '         mCol(CurTrack).KeySign = Data1
                End If
                Data2 = MFBuffer(MFFilePos) : MFFilePos += 1
                'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(Data1), "00") & " " & VB6.Format(Hex(Data2), "00"), TAB, MFFilePos - 2, TAB, MFFilePos, TAB, "Meta Event 59 Track = " & trnr))
                'Delta time juste apres
                l = GetVarLen()
                '      CurTrackTickPos = CurTrackTickPos + GetVarLen
            Case &H7FS
                l = GetVarLen()
                For X = 0 To (l - 1)
                    c = MFBuffer(MFFilePos) : MFFilePos += 1
                    Str_Renamed = Str_Renamed & Chr(MFBuffer(MFFilePos - 1))
                    'If DBUG Then System.Diagnostics.Debug.WriteLine(VB6.TabLayout(VB6.Format(Hex(c), "00"), TAB, MFFilePos - 1, TAB, MFFilePos, TAB, "Meta Event 7F Track = " & trnr & " ...Byte = " & Chr(MFBuffer(MFFilePos - 1))))
                Next
                l = GetVarLen()
            Case Else
                l = GetVarLen()
                For X = 0 To (l - 1)
                    c = MFBuffer(MFFilePos) : MFFilePos += 1
                    '         Str = Str & Chr(MFBuffer(MFFilePos - 1))
                    '         If DBUG Then Debug.Print Format(Hex$(c), "00"); Tab; MFFilePos - 1; Tab; MFFilePos; Tab; "Meta Event 7F Track = " & trnr & " ...Byte = " & Chr(MFBuffer(MFFilePos - 1))
                Next
                l = GetVarLen()
        End Select
        'running_status = &HFE ' &annule le running status
    End Sub

    '***************************************************************************************
    ' SysexEvent
    '***************************************************************************************
    Sub SysexEvent()
        Dim c As Byte
        Dim X As Integer
        Dim l As Integer

        ' If DBUG Then Debug.Print Format(Hex$(MFBuffer(MFFilePos - 1)), "00"); Tab; MFFilePos - 1; Tab; Filepos; Tab; "SysEx ... "
        ' Longueur du Sysex
        l = GetVarLen()
        For X = 0 To (l - 1)
            c = MFBuffer(MFFilePos) : MFFilePos += 1
            'Str = Str & Chr(MFBuffer(MFFilepos - 1))
            '    If DBUG Then Debug.Print Format(Hex$(c), "00"); Tab; Filepos - 1; Tab; Filepos; Tab; "SysEx ...Chr = " & Format(Hex$(c), "00")
        Next
        ' Delta Time
        l = GetVarLen()
    End Sub


End Class