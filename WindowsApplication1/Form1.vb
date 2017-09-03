Imports System.Management

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objWMI As New clsWMI()
        With objWMI
            cn.Text = .ComputerName
            cm.Text = .Manufacturer
            cmdl.Text = .Model

            Dim name As String
            name = .OsName
            Dim words As String() = name.Split(New Char() {"|"c})
            osn.Text = words.First

            osv.Text = .OSVersion
            st.Text = .SystemType
            tpm.Text = GetFileSize(CDbl(.TotalPhysicalMemory))
            wd.Text = .WindowsDirectory
            dcc.Text = GetFileSize(CDbl(.Capacityc))
            dfc.Text = GetFileSize(CDbl(.FreeSpacec))
            dcd.Text = GetFileSize(CDbl(.Capacityd))
            dfd.Text = GetFileSize(CDbl(.FreeSpaced))

            Dim capacity As Double
            Dim freespace As Double
            Dim persen As Integer

            capacity = .Capacityc
            freespace = .FreeSpacec
            persen = (freespace / capacity) * 100

            'dp.Text = Math.Round(persen) & "%"
        End With
    End Sub

    Dim DoubleBytes As Double

    Public Function GetFileSize(ByVal size As Double) As String        
        Dim TheSize As ULong = size

        Try
            Select Case TheSize
                Case Is >= 1099511627776
                    DoubleBytes = CDbl(TheSize / 1099511627776) 'TB
                    Return FormatNumber(DoubleBytes, 2) & " TB"
                Case 1073741824 To 1099511627775
                    DoubleBytes = CDbl(TheSize / 1073741824) 'GB
                    Return FormatNumber(DoubleBytes, 2) & " GB"
                Case 1048576 To 1073741823
                    DoubleBytes = CDbl(TheSize / 1048576) 'MB
                    Return FormatNumber(DoubleBytes, 2) & " MB"
                Case 1024 To 1048575
                    DoubleBytes = CDbl(TheSize / 1024) 'KB
                    Return FormatNumber(DoubleBytes, 2) & " KB"
                Case 0 To 1023
                    DoubleBytes = TheSize ' bytes
                    Return FormatNumber(DoubleBytes, 2) & " bytes"
                Case Else
                    Return ""
            End Select
        Catch
            Return ""
        End Try
    End Function

End Class
