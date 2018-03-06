Imports System.Runtime.InteropServices

Public Class Form1
    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    Private Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
    Private Const DBT_DEVTYP_VOLUME As Integer = &H2


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

    'Function that gets the drive letter from the unit mask
    Private Function GetDriveLetterFromMask(ByRef Unit As Int32) As Char
        For i As Integer = 0 To 25
            If Unit = (2 ^ i) Then
                Return Chr(Asc("A") + i)
            End If
        Next
    End Function

    'Override message processing to check for the DEVICECHANGE message
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        If m.Msg = WM_DEVICECHANGE Then
            If CInt(m.WParam) = DBT_DEVICEREMOVECOMPLETE Then
                Dim Volume As DEV_BROADCAST_VOLUME
                Volume = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                Dim DriveLetter As String = (GetDriveLetterFromMask(Volume.dbcv_unitmask) & ":\")
                MessageBox.Show("Drive removed " & DriveLetter)
            End If
            If CInt(m.WParam) = DBT_DEVICEARRIVAL Then
                Dim DeviceInfo As DEV_BROADCAST_HDR
                DeviceInfo = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_HDR)), DEV_BROADCAST_HDR)
                If DeviceInfo.dbch_devicetype = DBT_DEVTYP_VOLUME Then
                    Dim Volume As DEV_BROADCAST_VOLUME
                    Volume = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                    Dim DriveLetter As String = (GetDriveLetterFromMask(Volume.dbcv_unitmask) & ":\")
                    If IO.File.Exists(IO.Path.Combine(DriveLetter, "test.txt")) Then
                        '<<<< The test file has been found >>>>
                        MessageBox.Show("Found test file in " & DriveLetter)
                    Else
                        '<<<< Test file has not been found >>>>
                        MessageBox.Show("Could not find test file")
                    End If
                End If
            End If
        End If
        'Process all other messages as normal
        MyBase.WndProc(m)
    End Sub

End Class
