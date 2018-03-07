Imports System.IO
''' <summary>
''' This is an example class which can be used to manage the application My drive access and
''' anyother processes required
''' </summary>
Public Class MyDrive

    Private pMyDrive As String   ' MyDrive register

    ''' <summary>
    ''' Initialise MyDrive register
    ''' </summary>
    Public Sub New()
        pMyDrive = ""
        LookFoMyDrive()
    End Sub

#Region "Class properties"
    ReadOnly Property ApplicationDrive() As String
        Get
            Return pMyDrive
        End Get
    End Property
#End Region
#Region "Class public functions"
    ''' <summary>
    ''' Write a message on a windows form control
    ''' </summary>
    ''' <param name="frmtextbox"></param>
    Public Sub WriteMessage(ByRef frmtextbox As TextBox)
        If pMyDrive = "" Then
            frmtextbox.Text = "Application drive not available"
        Else
            frmtextbox.Text = "Application drive available " & pMyDrive
        End If
    End Sub
#End Region

#Region "Handel drive insert and remove events, which are passedto the top level form by OS"
    ''' <summary>
    ''' If there is no registered MyDrive test the inserted drive for markers
    ''' </summary>
    ''' <param name="DriveLetter"></param>
    Public Sub DriveInserted(DriveLetter As String)
        If pMyDrive = "" And IsItMarked(DriveLetter) Then pMyDrive = DriveLetter
    End Sub

    ''' <summary>
    ''' If removed drive is the  registered MyDrive then de-register it
    ''' </summary>
    ''' <param name="DriveLetter"></param>
    Public Sub DriveRemoved(DriveLetter As String)
        If pMyDrive = DriveLetter Then
            pMyDrive = ""
        End If
    End Sub
#End Region

#Region "Class private support functions"
    ''' <summary>
    ''' Test DriveLetter for the markers to determine if the drive inserted is MyDrive
    ''' 
    ''' In this case the drive ismarked with a file called "PhotoArchive.txt", in the top levelfolder
    ''' </summary>
    ''' <param name="DriveLetter"></param>
    Private Function IsItMarked(DriveLetter As String) As Boolean
        If IO.File.Exists(IO.Path.Combine(DriveLetter, "PhotoArchive.txt")) Then Return True
        Return False
    End Function

    ''' <summary>
    ''' Iterate the drives attached to this computer  to find one marked as this applications drive
    ''' </summary>
    Private Sub LookFoMyDrive()
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim d As DriveInfo

        For Each d In allDrives
            If d.IsReady = True Then
                If IsItMarked(d.Name) Then
                    pMyDrive = d.Name
                End If
            End If
        Next
    End Sub
#End Region
End Class
