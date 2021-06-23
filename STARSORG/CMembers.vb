Imports System.Data.SqlClient
Public Class CMembers
    'Represents the Member table and its associated business rules
    Private _Member As CMember
    'constructor
    Public Sub New()
        'instantiate the CMember object
        _Member = New CMember
    End Sub

    Public ReadOnly Property CurrentObject() As CMember
        Get
            Return _Member
        End Get
    End Property
    Public Sub Clear()
        _Member = New CMember
    End Sub
    Public Sub CreateNewRole()
        'call this when clearing the edit portion of the screen to add a new user
        Clear()
        _Member.IsNewMember = True
    End Sub
    Public Function Save() As Integer
        Return _Member.Save()
    End Function
    Public Function GetMemberByPID(strID As String) As CMember
        Dim params As New ArrayList
        params.Add(New SqlParameter("PID", strID))
        FillOBject(myDB.GetDataReaderBySP("sp_getMemberbyPID", params))
        Return _Member
    End Function
    Public Function GetAllMembers() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getMembers", Nothing)
        Return objDR
    End Function
    Private Function FillOBject(objDR As SqlDataReader) As CMember
        If objDR.Read Then
            With _Member
                .PID = objDR.Item("PID")
                .FName = objDR.Item("FName")
                .LName = objDR.Item("LName")
                .MI = objDR.Item("MI")
                .Phone = objDR.Item("Phone")
                .Email = objDR.Item("Email")
                .PhotoPath = objDR.Item("PhotoPath")

            End With
        Else 'no record was returned
            'nothing to do here
        End If
        objDR.Close()
        Return _Member
    End Function
End Class