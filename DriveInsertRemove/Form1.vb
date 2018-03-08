Imports System.Runtime.InteropServices

Public Class Form1
    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    Private Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
    Private Const DBT_DEVTYP_VOLUME As Integer = &H2

    Dim cMyDrive As New MyDrive                             ' A class to manage a USB drive dedicated to this application

    'Device information structure
    Public Structure DEV_BROADCAST_HDR
        Public dbch_size As Int32
        Public dbch_devicetype As Int32
        Public dbch_reserved As Int32
    End Structure

    'Volume information Structure
    Private Structure DEV_BROADCAST_VOLUME
        Public dbcv_size As Int32
        Public dbcv_devicetype As Int32
        Public dbcv_reserved As Int32
        Public dbcv_unitmask As Int32
        Public dbcv_flags As Int16
    End Structure

    ''' <summary>
    ''' Get the drive letter from the unit mask
    ''' </summary>
    ''' <param name="Unit"></param>
    ''' <returns></returns>
    Private Function GetDriveLetterFromMask(ByRef Unit As Int32) As Char
        GetDriveLetterFromMask = ""
        For i As Integer = 0 To 25
            If Unit = (2 ^ i) Then
                GetDriveLetterFromMask = Chr(Asc("A") + i)
            End If
        Next
    End Function


    ''' <summary>
    ''' Override message processing to check for the DEVICECHANGE message
    ''' 
    ''' This was modified slightly from code found on web sorry dont know attribution any more
    ''' I tried lots of stuff, this worked for me and is simple
    ''' </summary>
    ''' <param name="m"></param>
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        If m.Msg = WM_DEVICECHANGE Then
            If CInt(m.WParam) = DBT_DEVICEREMOVECOMPLETE Then
                Dim Volume As DEV_BROADCAST_VOLUME
                Volume = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                cMyDrive.DriveRemoved((GetDriveLetterFromMask(Volume.dbcv_unitmask) & ":\"))
                cMyDrive.WriteMessage(Me.TextBox1)
            End If
            If CInt(m.WParam) = DBT_DEVICEARRIVAL Then
                Dim DeviceInfo As DEV_BROADCAST_HDR
                DeviceInfo = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_HDR)), DEV_BROADCAST_HDR)
                If DeviceInfo.dbch_devicetype = DBT_DEVTYP_VOLUME Then
                    Dim Volume As DEV_BROADCAST_VOLUME
                    Volume = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                    'Dim DriveLetter As String = (GetDriveLetterFromMask(Volume.dbcv_unitmask) & ":\")
                    cMyDrive.DriveInserted((GetDriveLetterFromMask(Volume.dbcv_unitmask) & ":\"))
                    cMyDrive.WriteMessage(Me.TextBox1)
                End If
            End If
        End If
        'Process all other messages as normal
        MyBase.WndProc(m)
    End Sub


    ''' <summary>
    ''' Just in case the application drive is attached before the application starts
    ''' Do your stuff on load the application drive class checks for the drive when
    ''' it itialises
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        cMyDrive.WriteMessage(Me.TextBox1)
        ' ----------------  put other on load stuff here
    End Sub

    ' ----------------  put other application stuff here

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        cMyDrive.WriteMessage(Me.TextBox1)
    End Sub
End Class
