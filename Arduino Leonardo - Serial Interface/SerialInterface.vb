Imports System.Runtime.InteropServices
Imports System.IO.Ports
Imports System.Management

Module SerialInterface
    'Change the boardName based on the Board Used
    Private boardName As String = "Arduino Leonardo"
    Private portName As String


    Sub Main()
        'Intialize Active Port Names in to ports
        Dim ports As String() = SerialPort.GetPortNames

        Try
            'Query Details of Devices connected
            Dim searcher As New ManagementObjectSearcher( _
            "root\cimv2", _
            "SELECT * FROM Win32_SerialPort")

            For Each queryObj As ManagementObject In searcher.Get()
                Dim comName As String
                Dim arduinoExists As Integer

                'Retrieve Name of the Devices
                comName = queryObj("Name")

                'Check if the name consists boardName in it
                arduinoExists = comName.IndexOf(boardName)

                'arduinoExists is set to -1 if the IndexOf Method doesnt fine boardName
                If Not arduinoExists = -1 Then
                    'Retrieve Port Name
                    Dim position As Integer = comName.IndexOf("COM")
                    Dim comNameLength As Integer
                    Dim temp As String = comName.Substring(position + 4, 1)

                    If temp = "0" Then
                        comNameLength = 5
                    End If

                    portName = comName.Substring(position, comNameLength)
                    Console.WriteLine("Device Found on " & portName)
                Else
                    Console.WriteLine("Device Not Found")
                    Console.WriteLine("Press Any Key To Exit")
                    Console.ReadLine()
                    Exit Sub
                End If

            Next
        Catch err As ManagementException
            Console.WriteLine("An error occurred while querying for WMI data: " & err.Message)
            Console.WriteLine("Press Any Key To Exit")
            Console.ReadLine()
        End Try

        'Connecting Port to Read and Write Data
        Dim port As New System.IO.Ports.SerialPort
        Dim fetchedData

        'Intialize Port Settings
        port.PortName = portName

        'Open Port
        port.Open()
        port.BaudRate = 9600
        port.DtrEnable = True
        port.RtsEnable = True

        'Loops without exit for receiving and Writing Data
        While (True)
            'Read Incoming Data
            fetchedData = port.ReadByte
            Console.WriteLine(fetchedData)
            'WriteData
            port.Write("Hello World!")
        End While

    End Sub


End Module
