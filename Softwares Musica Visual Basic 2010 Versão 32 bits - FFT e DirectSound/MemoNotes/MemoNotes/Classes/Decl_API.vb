Option Explicit On
Imports System.Runtime.InteropServices

Public Structure MIDIOUTCAPS
    Dim ManufacturerID As Short
    Dim ProductID As Short
    Dim DriverVersion As Integer
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=64)> Dim ProductName As Byte()
    Dim Technology As Short
    Dim Voices As Short
    Dim Notes As Short
    Dim ChannelMask As Short
    Dim Support As Integer
End Structure

Public Structure MIDIINCAPS
    Dim ManufacturerID As Short
    Dim ProductID As Short
    Dim DriverVersion As Integer
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=64)> Dim ProductName As Byte()
End Structure


Public Structure TIMECAPS
    Public PeriodMin As Integer
    Public PeriodMax As Integer
End Structure

Public Delegate Sub TimeProc(ByVal Id As Integer, ByVal Msg As Integer, ByVal User As Integer, ByVal Param1 As Integer, ByVal Param2 As Integer)
Public Delegate Sub MidiDelegate(ByVal MidiInHandle As Int32, ByVal wMsg As Int32, ByVal Instance As Int32, ByVal wParam As Int32, ByVal lParam As Int32)

Public NotInheritable Class WinMM
    '*** MIDI OUT ***
    Public Declare Function midiOutGetNumDevs _
                   Lib "winmm.dll" () _
                   As Integer

    Public Declare Auto Function midiOutGetDevCaps _
                   Lib "winmm.dll" ( _
                   ByVal DevNum As Integer, _
                   ByRef DevCaps As MIDIOUTCAPS, _
                   ByVal SizeOfStruc As Integer) _
                   As Integer

    Public Declare Function midiOutOpen _
                   Lib "winmm.dll" ( _
                   ByRef hMidiOut As Integer, _
                   ByVal devID As Integer, _
                   ByVal cbfunc As Integer, _
                   ByVal cbdata As Integer, _
                   ByVal cboptions As Integer) _
                   As Integer

    Public Declare Sub midiOutClose _
                   Lib "winmm.dll" ( _
                   ByVal hMidiOut _
                   As Integer)

    Public Declare Function midiOutReset _
                   Lib "winmm.dll" ( _
                   ByVal hMidiOut As Integer) _
                   As Integer

    Public Declare Function midiOutShortMsg _
                   Lib "winmm.dll" ( _
                   ByVal hMidiOut As Integer, _
                   ByVal msg As Integer) _
                   As Integer

    '*** MIDI IN ***
    Public Declare Function midiInGetNumDevs _
               Lib "winmm.dll" () _
               As Integer

    Public Declare Auto Function midiInGetDevCaps _
                   Lib "winmm.dll" ( _
                   ByVal DevNum As Integer, _
                   ByRef DevCaps As MIDIINCAPS, _
                   ByVal SizeOfStruc As Integer) _
                   As Integer

    Public Declare Function midiInOpen _
                   Lib "winmm.dll" ( _
                   ByRef lphMidiIn As Integer, _
                   ByVal uDeviceID As Integer, _
                   <MarshalAs(UnmanagedType.FunctionPtr)> ByVal dwCallback As MidiDelegate, _
                   ByVal dwInstance As Integer, _
                   ByVal dwFlags As Integer) As Integer

    Public Declare Function midiInClose _
                   Lib "winmm.dll" ( _
                   ByVal hMidiIn As Integer) As Integer

    Public Declare Function midiInStart _
                   Lib "winmm.dll" ( _
                   ByVal hMidiIn As Integer) As Integer

    Public Declare Function midiInStop _
                   Lib "winmm.dll" ( _
                   ByVal hMidiIn As Integer) As Integer

    Public Declare Function midiInReset _
                   Lib "winmm.dll" ( _
                   ByVal hMidiIn As Integer) As Integer


    '*** Timer ***
    Public Declare Function timeBeginPeriod _
                   Lib "winmm.dll" ( _
                   ByVal uPeriod As Integer) _
                   As Integer

    Public Declare Function timeEndPeriod _
                   Lib "winmm.dll" ( _
                   ByVal uPeriod As Integer) _
                   As Integer

    Public Declare Function timeSetEvent _
                   Lib "winmm.dll" ( _
                   ByVal Delay As Integer, _
                   ByVal Resolution As Integer, _
                   ByVal CBFunction As TimeProc, _
                   ByVal User As Integer, _
                   ByVal Flags As Integer) _
                   As Integer

    Public Declare Function timeKillEvent _
                   Lib "winmm.dll" ( _
                   ByVal uID As Integer) _
                   As Integer

    Public Declare Function timeGetDevCaps _
                   Lib "winmm.dll" ( _
                   ByRef lpTimeCaps As TIMECAPS, _
                   ByVal uSize As Integer) _
                   As Integer

End Class

Public NotInheritable Class Win32a
    '*** Query Performance ***
    Public Declare Auto Function QueryPerformanceCounter _
                   Lib "kernel32.dll" ( _
                   ByRef Counter As Long) _
                   As Long

    Public Declare Auto Function QueryPerformanceFrequency _
                   Lib "kernel32.dll" ( _
                   ByRef counter As Long) _
                   As Long

End Class