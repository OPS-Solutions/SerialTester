Imports System.IO.Ports.SerialPort
Imports System.IO.Ports
Imports System.IO

Public Class Form1
    Private NEC As Byte() = {3, &H96, 0, 0, 2, 0, 1, &H9C}
    Private ReadBuffer(11) As Byte
    Private strPorts As String() = {""}
    Private strPort As String = ""

    Private mintBaud As Integer = 38400
    Private mintBits As Integer = 8
    Private menmStop As StopBits = Ports.StopBits.One
    Private menmParity As Parity = Ports.Parity.None
    Private menmFlowControl As IO.Ports.Handshake = Ports.Handshake.None

    Private mobjSerial As SerialPort = Nothing

    Private strNewData As String = ""


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckCOMPorts()
    End Sub

    Private Function CheckCOMPorts() As Boolean
        Try
            strPorts = SerialPort.GetPortNames

            NumericUpDown1.Value = CInt(strPorts.Last.Last.ToString)
            strPort = strPorts.Last

            ' Look for COM ports
            If (strPorts Is Nothing) OrElse (strPorts.Length = 0) Then
                ' No COM ports at all
              
                Return False
            ElseIf (Not strPorts.Contains(strPort)) Then
                ' Particular COM port is missing
                Dim strMessage As String = strPorts(0)
                If (UBound(strPorts) > 0) Then
                    For i As Integer = 1 To UBound(strPorts)
                        strMessage &= ", " & strPorts(i)
                    Next
                End If
             
                Return False
            Else
                ' Proper port is found
                Return True
            End If


        Catch exUA As UnauthorizedAccessException
    

        Catch exTh As Threading.ThreadAbortException
            ' Do nothing, aborting thread

        Catch ex As Exception
           
        End Try

        Return False
    End Function

    Private Function OpenCOMPort(ByRef objSerial As SerialPort) As Boolean

        Try

            StopCommunications(objSerial)

            objSerial = New SerialPort(strPort, Me.mintBaud, Me.menmParity, Me.mintBits, Me.menmStop)
            objSerial.Handshake = Me.menmFlowControl
            ClosePort(objSerial)


            objSerial.Open()
            Return True
        Catch exUA As UnauthorizedAccessException


        Catch exTh As Threading.ThreadAbortException
            ' Do nothing, aborting thread

        Catch ex As Exception

        End Try

        Return False
    End Function

    Private Sub StopCommunications(ByRef objSerial As SerialPort)
        ClosePort(objSerial)

        Try
            If (objSerial IsNot Nothing) Then
                objSerial.Dispose()
                objSerial = Nothing
            End If
        Catch exTh As Threading.ThreadAbortException
            ' Do nothing, aborting thread
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ClosePort(ByRef objSerial As SerialPort)
        Try
            If (objSerial IsNot Nothing) AndAlso (objSerial.IsOpen) Then
                objSerial.Close()
            End If
        Catch exTh As Threading.ThreadAbortException
            ' Do nothing, aborting thread
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim mythread As System.Threading.Thread

        mythread = New Threading.Thread(AddressOf threadstuff)
        mythread.Priority = Threading.ThreadPriority.Normal
        mythread.IsBackground = True
        mythread.Name = Format(Now, "YYYY-dd-MM hh:mm:ss tt")
        mythread.Start()

    End Sub


    Public Event UpdateText(message As String)

    Private Delegate Sub eventhandlerdelegate(message As String)


    Private Sub eventhandler(message As String) Handles Me.UpdateText
        Static Dim toggle As Boolean = False

        If Not Me.InvokeRequired Then

            TextBox1.Text = message & TextBox1.Text

            If toggle Then
                Me.BackColor = System.Drawing.Color.Blue
            Else
                Me.BackColor = System.Drawing.Color.Yellow
            End If

            toggle = Not toggle

            Me.Update()
            Me.Refresh()
        Else
            Dim objDelegate As New eventhandlerdelegate(AddressOf eventhandler)
            Me.BeginInvoke(objDelegate, New Object() {message})

        End If
    End Sub

    Private Sub threadstuff()
        '  mobjSerial = New SerialPort(strPort, mintBaud, menmParity, mintBits, menmStop)
        OpenCOMPort(mobjSerial)

        Dim count As Integer = 0
        mobjSerial.ReadTimeout = 500


        While True
            Try
                mobjSerial.Write(NEC, 0, 8)
                System.Threading.Thread.Sleep(100)
                mobjSerial.Read(ReadBuffer, 0, 12)  ' Read latest data

                count += 1

                Dim newtext As String = ""


                For i As Integer = 0 To ReadBuffer.Length - 1
                    newtext &= "[" & ReadBuffer(i).ToString & "]"
                Next


                newtext = vbCrLf & count.ToString & "-" & newtext
                RaiseEvent UpdateText(newtext)


                mobjSerial.DiscardInBuffer() ' This is important
                mobjSerial.Close()           ' This is also important
                mobjSerial.Open()            ' yes, this too
            Catch ex As Exception

            End Try

        End While
    End Sub
End Class
