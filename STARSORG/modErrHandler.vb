Module modErrHandler
    'useful routines for input validation
    Public Function ValidateTextBoxLength(ByRef obj As TextBox, ByRef errP As ErrorProvider) As Boolean
        'This procedure validates that a textbox is not empty
        If obj.Text.Length = 0 Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a value here")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function
    Public Function ValidateTextBoxNumeric(ByRef obj As TextBox, ByRef errP As ErrorProvider) As Boolean
        'This procedure validates that a textbox has a numeric value
        If Not IsNumeric(obj.Text) Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a numeric value here")
            obj.Focus()
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function
    Public Function ValidateTextBoxDate(ByRef obj As TextBox, ByRef errP As ErrorProvider) As Boolean
        'This procedure validates that textbox has a valid date value
        If Not IsDate(obj.Text) Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a valid date here")
            obj.Focus()
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function

    Public Function ValidateCombo(ByRef obj As ComboBox, ByRef errP As ErrorProvider) As Boolean
        If obj.SelectedIndex = -1 Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must make a selection here")
            obj.Focus()
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function

    Public Function ValidateMaskedTextBoxLength(ByRef obj As MaskedTextBox, ByRef errP As ErrorProvider) As Boolean
        If obj.Text.Length = 0 Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a value here")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function
End Module
