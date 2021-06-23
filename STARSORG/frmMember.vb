Imports System.Data.SqlClient
Imports System.IO
Public Class frmMember
    Private sqlDA As SqlDataAdapter
    Private dt As DataTable
    Private myDB As CDB
    Private sqlDR As SqlDataReader
    Private objMembers As CMembers
    Private blnReloading As Boolean
    Private blnClearing As Boolean
    Private strPhotoPath As String
#Region "Toolbar routines"
    Private Sub tsbCourse_Click(sender As Object, e As EventArgs) Handles tsbCourse.Click
        intNextAction = ACTION_COURSE
        Me.Hide()
    End Sub
    Private Sub tsbEvent_Click(sender As Object, e As EventArgs) Handles tsbEvent.Click
        intNextAction = ACTION_EVENT
        Me.Hide()
    End Sub
    Private Sub tsbHelp_Click(sender As Object, e As EventArgs) Handles tsbHelp.Click
        intNextAction = ACTION_HELP
        Me.Hide()
    End Sub
    Private Sub tsbHome_Click(sender As Object, e As EventArgs) Handles tsbHome.Click
        intNextAction = ACTION_HOME
        Me.Hide()
    End Sub
    Private Sub tsbLogout_Click(sender As Object, e As EventArgs) Handles tsbLogOut.Click
        intNextAction = ACTION_LOGOUT
        Me.Hide()
    End Sub
    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        'do nothing
    End Sub
    Private Sub tsbRole_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        intNextAction = ACTION_ROLE
        Me.Hide()
    End Sub
    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs) Handles tsbRSVP.Click
        intNextAction = ACTION_RSVP
        Me.Hide()
    End Sub
    Private Sub tsbsemester_Click(sender As Object, e As EventArgs) Handles tsbSemester.Click
        intNextAction = ACTION_SEMESTER
        Me.Hide()
    End Sub
    Private Sub tsbtutor_Click(sender As Object, e As EventArgs) Handles tsbTutor.Click
        intNextAction = ACTION_TUTOR
        Me.Hide()
    End Sub
    Private Sub tsbAdmin_Click(sender As Object, e As EventArgs) Handles tsbAdmin.Click
        intNextAction = ACTION_ADMIN
        Me.Hide()
    End Sub
    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs) Handles tsbCourse.MouseEnter, tsbEvent.MouseEnter, tsbHelp.MouseEnter, tsbHome.MouseEnter, tsbLogOut.MouseEnter, tsbMember.MouseEnter, tsbRole.MouseEnter, tsbRSVP.MouseEnter, tsbSemester.MouseEnter, tsbTutor.MouseEnter, tsbAdmin.MouseEnter
        'We need to do this only because we are not putting our image ein the image property of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text
    End Sub

    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs) Handles tsbCourse.MouseLeave, tsbEvent.MouseLeave, tsbHelp.MouseLeave, tsbHome.MouseLeave, tsbLogOut.MouseLeave, tsbMember.MouseLeave, tsbRole.MouseLeave, tsbRSVP.MouseLeave, tsbSemester.MouseLeave, tsbTutor.MouseLeave, tsbAdmin.MouseLeave
        'We need to do this only because we are not putting our image ein the image property of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub
#End Region
    Private Sub frmMember_Load(sender As Object, e As EventArgs) Handles Me.Load
        myDB = New CDB

        If Not myDB.OpenDB Then
            Application.Exit()
        End If


        objMembers = New CMembers
        LoadMembers()
        grpInformation.Enabled = False

    End Sub
    Private Function CalcLocation(grpbox As GroupBox, subForm As UserControl) As Point
        Return New Point((grpbox.Width - subForm.Width) / 2, (grpbox.Height - subForm.Height) / 2)
    End Function
    Private Sub LoadMembers()

        Dim objDR As SqlDataReader
        lstMembers.Items.Clear()

        Try
            objDR = objMembers.GetAllMembers()
            Do While objDR.Read
                lstMembers.Items.Add(objDR.Item("PID") & " " & objDR.Item("LName") & ", " & objDR.Item("Fname"))
            Loop
            objDR.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        If objMembers.CurrentObject.PID <> "" Then
            lstMembers.SelectedIndex = lstMembers.FindStringExact(objMembers.CurrentObject.PID)
        End If
        blnReloading = False
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim blnErrors As Boolean
        Dim params As New ArrayList
        Dim objDR As SqlDataReader
        lstMembers.Items.Clear()
        params.Add(New SqlParameter("@lastname", txtSearchBox.Text))
        objDR = myDB.GetDataReaderBySP("dbo.sp_getMemberName", params)
        Do While objDR.Read
            lstMembers.Items.Add(objDR.Item("PID") & " " & objDR.Item("LName") & ", " & objDR.Item("Fname"))
        Loop
        objDR.Close()
    End Sub

    Private Sub LoadSelectedRecord()
        Dim ofdUpload As New OpenFileDialog
        Try
            objMembers.GetMemberByPID(lstMembers.SelectedItem.ToString)
            With objMembers.CurrentObject
                txtPID.Text = .PID
                txtFName.Text = .FName
                txtLName.Text = .LName
                txtMI.Text = .MI
                txtEmail.Text = .Email
                mskPhone.Text = .Phone
                picMember.ImageLocation = .PhotoPath
            End With
        Catch ex As Exception
            MessageBox.Show("Error loading Role values: " & ex.ToString, "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub lstMembers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMembers.SelectedIndexChanged
        If blnClearing Then
            Exit Sub
        End If
        If blnReloading Then
            ToolStrip.Text = ""
            Exit Sub
        End If
        If lstMembers.SelectedIndex = -1 Then 'nothing to do
            Exit Sub
        End If
        chkNewMem.Checked = False
        LoadSelectedRecord()
        grpInformation.Enabled = True
    End Sub

    Private Sub chkNewMem_CheckedChanged(sender As Object, e As EventArgs) Handles chkNewMem.CheckedChanged
        If blnClearing Then
            Exit Sub
        End If
        If chkNewMem.Checked Then
            ToolStrip.Text = ""
            txtPID.Clear()
            txtFName.Clear()
            txtLName.Clear()
            txtMI.Clear()
            txtEmail.Clear()
            mskPhone.Clear()
            picMember.ImageLocation = ""

            lstMembers.SelectedIndex = -1
            grpMembers.Enabled = False
            grpSearch.Enabled = False
            grpInformation.Enabled = True
            objMembers.CreateNewRole()

            txtPID.Focus()
        Else
            grpInformation.Enabled = False
            grpSearch.Enabled = True
            grpMembers.Enabled = True
            grpInformation.Enabled = True
            objMembers.CurrentObject.IsNewMember = False
            picMember.ImageLocation = ""

        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim intResult As Integer
        Dim blnErrors As Boolean
        Dim ofdUpload As New OpenFileDialog
        blnErrors = False
        errP.Clear()
        ToolStrip.Text = ""
        '-------- add your validation here -------
        If Not ValidateTextBoxLength(txtPID, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtFName, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtLName, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtMI, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtEmail, errP) Then
            blnErrors = True
        End If
        If strPhotoPath = "" Then
            errP.SetError(btnUpload, "Please upload a picture")
            blnErrors = True
        End If
        If Not ValidateMaskedTextBoxLength(mskPhone, errP) Then
            blnErrors = True
        End If
        If blnErrors Then
            Exit Sub
        End If
        'if we get this far all the data is good
        With objMembers.CurrentObject
            .PID = txtPID.Text
            .FName = txtFName.Text
            .LName = txtLName.Text
            .MI = txtMI.Text
            .Email = txtEmail.Text
            .Phone = mskPhone.Text
            .PhotoPath = strPhotoPath

        End With

        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objMembers.Save
            If intResult = 1 Then
                MessageBox.Show("Member Saved ", "New Member", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            If intResult = -1 Then 'Role ID was not unique when adding a record
                MessageBox.Show(" PID must be unique. Unable to save a new Member.", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ToolStrip.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save new member: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ToolStrip.Text = "Error"
        End Try
        Me.Cursor = Cursors.Default
        blnReloading = True
        LoadMembers() 'reload so that a new saved record we appear in the list
        chkNew.Checked = False
        grpMembers.Enabled = True
        blnReloading = False


        txtPID.Clear()
        txtFName.Clear()
        txtLName.Clear()
        txtMI.Clear()
        txtEmail.Clear()
        mskPhone.Clear()
        picMember.ImageLocation = ""
    End Sub

    Private Sub btnUpload_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        Dim ofdUpload As New OpenFileDialog

        ofdUpload.Filter = "Picture Files (*)|* .bmp;* .gif;* .jpg"
        If ofdUpload.ShowDialog = Windows.Forms.DialogResult.OK Then
            picMember.Image = Image.FromFile(ofdUpload.FileName)
            strPhotoPath = ofdUpload.FileName
        End If
    End Sub



End Class