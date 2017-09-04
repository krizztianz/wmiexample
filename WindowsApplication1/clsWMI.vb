Imports System.Management

Public Class clsWMI
    Private objOS As ManagementObjectSearcher
    Private objCS As ManagementObjectSearcher
    Private objDiskc As ManagementObjectSearcher
    Private objDiskd As ManagementObjectSearcher
    Private objMgmt As ManagementObject
    Private m_strComputerName As String
    Private m_strManufacturer As String
    Private m_StrModel As String
    Private m_strOSName As String
    Private m_strOSVersion As String
    Private m_strSystemType As String
    Private m_strTPM As String
    Private m_strFPM As String
    Private m_strWindowsDir As String
    Private m_strCapacityc As String
    Private m_strFreeSpacec As String
    Private m_strCapacityd As String
    Private m_strFreeSpaced As String


    Public Sub New()

        objOS = New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
        objCS = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        objDiskc = New ManagementObjectSearcher("SELECT * FROM Win32_Volume WHERE DriveLetter = 'c:'")
        objDiskd = New ManagementObjectSearcher("SELECT * FROM Win32_Volume WHERE DriveLetter = 'd:'")
        For Each objMgmt In objOS.Get


            m_strOSName = objMgmt("name").ToString()
            m_strOSVersion = objMgmt("version").ToString()
            m_strComputerName = objMgmt("csname").ToString()
            m_strWindowsDir = objMgmt("windowsdirectory").ToString()
            m_strFPM = objMgmt("FreePhysicalMemory").ToString()
        Next

        For Each objMgmt In objCS.Get
            m_strManufacturer = objMgmt("manufacturer").ToString()
            m_StrModel = objMgmt("model").ToString()
            m_strSystemType = objMgmt("systemtype").ToString
            m_strTPM = objMgmt("totalphysicalmemory").ToString()
        Next

        For Each objMgmt In objDiskc.Get
            m_strCapacityc = objMgmt("Capacity").ToString()
            m_strFreeSpacec = objMgmt("FreeSpace").ToString()
        Next

        For Each objMgmt In objDiskd.Get
            m_strCapacityd = objMgmt("Capacity").ToString()
            m_strFreeSpaced = objMgmt("FreeSpace").ToString()
        Next
    End Sub

    Public ReadOnly Property ComputerName()
        Get
            ComputerName = m_strComputerName
        End Get

    End Property
    Public ReadOnly Property Manufacturer()
        Get
            Manufacturer = m_strManufacturer
        End Get

    End Property
    Public ReadOnly Property Model()
        Get
            Model = m_StrModel
        End Get

    End Property
    Public ReadOnly Property OsName()
        Get
            OsName = m_strOSName
        End Get

    End Property

    Public ReadOnly Property OSVersion()
        Get
            OSVersion = m_strOSVersion
        End Get

    End Property
    Public ReadOnly Property SystemType()
        Get
            SystemType = m_strSystemType
        End Get

    End Property
    Public ReadOnly Property TotalPhysicalMemory()
        Get
            TotalPhysicalMemory = m_strTPM
        End Get

    End Property
    Public ReadOnly Property FreePhysicalMemory()
        Get
            FreePhysicalMemory = (CDbl(m_strFPM) * 1024)
        End Get

    End Property

    Public ReadOnly Property WindowsDirectory()
        Get
            WindowsDirectory = m_strWindowsDir
        End Get

    End Property

    Public ReadOnly Property FreeSpacec()
        Get
            FreeSpacec = m_strFreeSpacec
        End Get

    End Property

    Public ReadOnly Property Capacityc()
        Get
            Capacityc = m_strCapacityc
        End Get

    End Property

    Public ReadOnly Property FreeSpaced()
        Get
            FreeSpaced = m_strFreeSpaced
        End Get

    End Property

    Public ReadOnly Property Capacityd()
        Get
            Capacityd = m_strCapacityd
        End Get

    End Property

End Class