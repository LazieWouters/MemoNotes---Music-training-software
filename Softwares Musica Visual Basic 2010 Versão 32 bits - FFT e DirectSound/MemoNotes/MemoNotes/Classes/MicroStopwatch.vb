Namespace MicroLibrary

    Public Class MicroStopwatch
        Inherits System.Diagnostics.Stopwatch
        Private m_dMicroSecPerTick As Double = 1000000.0 / Frequency

        Public Sub New()
            If Not System.Diagnostics.Stopwatch.IsHighResolution Then
                Throw New Exception("On this system the high-resolution performance counter is not available")
            End If
        End Sub

        Public ReadOnly Property ElapsedMicroseconds() As Long
            Get
                Return CLng(Math.Truncate(ElapsedTicks * m_dMicroSecPerTick))
            End Get
        End Property
    End Class

    Public Class MicroTimer
        Public Delegate Sub MicroTimerElapsedEventHandler(ByVal sender As Object, ByVal timerEventArgs As MicroTimerEventArgs)
        Public Event MicroTimerElapsed As MicroTimerElapsedEventHandler

        Private m_threadTimer As System.Threading.Thread = Nothing
        Private m_lIgnoreEventIfLateBy As Long = Long.MaxValue
        Private m_lTimerIntervalInMicroSec As Long = 0
        Private m_bStopTimer As Boolean = True

        Public Sub New()
        End Sub

        Public Sub New(ByVal lTimerIntervalInMicroseconds As Long)
            Interval = lTimerIntervalInMicroseconds
        End Sub

        Public Property Interval() As Long
            Get
                Return m_lTimerIntervalInMicroSec
            End Get
            Set(ByVal value As Long)
                m_lTimerIntervalInMicroSec = value
            End Set
        End Property

        Public Property IgnoreEventIfLateBy() As Long
            Get
                Return m_lIgnoreEventIfLateBy
            End Get
            Set(ByVal value As Long)
                If value = 0 Then
                    m_lIgnoreEventIfLateBy = Long.MaxValue
                Else
                    m_lIgnoreEventIfLateBy = value
                End If
            End Set
        End Property

        Public Property Enabled() As Boolean
            Get
                Return (m_threadTimer IsNot Nothing AndAlso m_threadTimer.IsAlive)
            End Get
            Set(ByVal value As Boolean)
                If value Then
                    Start()
                Else
                    [Stop]()
                End If
            End Set
        End Property

        Public Sub Start()
            If (m_threadTimer Is Nothing OrElse Not m_threadTimer.IsAlive) AndAlso Interval > 0 Then
                m_bStopTimer = False
                Dim threadStart As System.Threading.ThreadStart = Function()
                                                                      NotificationTimer(Interval, IgnoreEventIfLateBy, m_bStopTimer)
                                                                  End Function
                m_threadTimer = New System.Threading.Thread(threadStart)
                m_threadTimer.Priority = System.Threading.ThreadPriority.Highest
                m_threadTimer.Start()
            End If
        End Sub

        Public Sub [Stop]()
            m_bStopTimer = True

            While Enabled
            End While
        End Sub

        Private Sub NotificationTimer(ByVal lTimerInterval As Long, ByVal lIgnoreEventIfLateBy As Long, ByRef bStopTimer As Boolean)
            Dim nTimerCount As Integer = 0
            Dim lNextNotification As Long = 0
            Dim lCallbackFunctionExecutionTime As Long = 0

            Dim microStopwatch As New MicroStopwatch()
            microStopwatch.Start()

            While Not bStopTimer
                lCallbackFunctionExecutionTime = microStopwatch.ElapsedMicroseconds - lNextNotification
                lNextNotification += lTimerInterval
                nTimerCount += 1
                Dim lElapsedMicroseconds As Long = 0

                While (InlineAssignHelper(lElapsedMicroseconds, microStopwatch.ElapsedMicroseconds)) < lNextNotification
                End While

                Dim lTimerLateBy As Long = lElapsedMicroseconds - (nTimerCount * lTimerInterval)

                If lTimerLateBy < lIgnoreEventIfLateBy Then
                    Dim microTimerEventArgs As New MicroTimerEventArgs(nTimerCount, lElapsedMicroseconds, lTimerLateBy, lCallbackFunctionExecutionTime)
                    RaiseEvent MicroTimerElapsed(Me, microTimerEventArgs)
                End If
            End While

            microStopwatch.[Stop]()
        End Sub
        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
            target = value
            Return value
        End Function
    End Class

    Public Class MicroTimerEventArgs
        Inherits EventArgs
        Public Property TimerCount() As Integer
            Get
                Return m_TimerCount
            End Get
            Private Set(ByVal value As Integer)
                m_TimerCount = Value
            End Set
        End Property
        Private m_TimerCount As Integer
        ' Simple counter, number times timed event (callback function) executed
        Public Property ElapsedMicroseconds() As Long
            Get
                Return m_ElapsedMicroseconds
            End Get
            Private Set(ByVal value As Long)
                m_ElapsedMicroseconds = Value
            End Set
        End Property
        Private m_ElapsedMicroseconds As Long
        ' Time when timed event was called since timer started
        Public Property TimerLateBy() As Long
            Get
                Return m_TimerLateBy
            End Get
            Private Set(ByVal value As Long)
                m_TimerLateBy = Value
            End Set
        End Property
        Private m_TimerLateBy As Long
        ' How late the timer was compared to when it should have been called
        Public Property CallbackFunctionExecutionTime() As Long
            Get
                Return m_CallbackFunctionExecutionTime
            End Get
            Private Set(ByVal value As Long)
                m_CallbackFunctionExecutionTime = Value
            End Set
        End Property
        Private m_CallbackFunctionExecutionTime As Long
        ' The time it took to execute the previous call to the callback function (OnTimedEvent)
        Public Sub New(ByVal nTimerCount As Integer, ByVal lElapsedMicroseconds As Long, ByVal lTimerLateBy As Long, ByVal lCallbackFunctionExecutionTime As Long)
            TimerCount = nTimerCount
            ElapsedMicroseconds = lElapsedMicroseconds
            TimerLateBy = lTimerLateBy
            CallbackFunctionExecutionTime = lCallbackFunctionExecutionTime
        End Sub
    End Class
End Namespace

