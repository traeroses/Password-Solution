' Name:         Password Project
' Purpose:      Create a password.
' Programmer:   <your name> on <current date>

Option Explicit Off
Option Strict Off
Option Infer On

Public Class frmMain
    Dim dblL As Integer
    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        ' Create a password.
        Dim strWords As String
        Dim strPassword As String
        Dim intSpaceIndex As Integer
        Dim strFinal As String
        Dim strFirst As String
        strWords = txtWords.Text.Trim
        If strWords <> String.Empty Then

            'Assign the first character as the password'
            strPassword = strWords(0)
            'Search for the first space in the input'
            intSpaceIndex = strWords.IndexOf(" ")
            Do Until intSpaceIndex = -1
                ' Concatenate the character that follows
                ' the space to the password.
                strPassword = strPassword & strWords(intSpaceIndex + 1)
                ' Search for the next space.
                intSpaceIndex = strWords.IndexOf(" ", intSpaceIndex + 1)
            Loop
            ' Insert the number after the first character.
            strPassword = strPassword.Insert(1, strPassword.Length.ToString)


            strFinal = strPassword
            strFirst = strFinal.ElementAt(0)
            If Asc(strFirst) > 64 AndAlso Asc(strFirst) < 91 Then
                dblL = 1
            ElseIf Asc(strFirst) > 96 AndAlso Asc(strFirst) < 123 Then
                dblL = 2
            End If

            For Each c As Char In strFinal
                If c = strFirst = False Then
                    If Asc(c) > 64 AndAlso Asc(c) < 91 Then
                        If dblL = 1 Then
                            strFinal = strFinal.Replace(c, c.ToString.ToLower)
                            dblL = 2
                        ElseIf dblL = 2 Then
                            strFinal = strFinal.Replace(c, c.ToString.ToUpper)
                            dblL = 1
                        End If
                    ElseIf Asc(c) > 96 AndAlso Asc(c) < 123 Then
                        If dblL = 1 Then
                            strFinal = strFinal.Replace(c, c.ToString.ToLower)
                            dblL = 2
                        ElseIf dblL = 2 Then
                            strFinal = strFinal.Replace(c, c.ToString.ToUpper)
                            dblL = 1
                        End If

                    End If
                End If
            Next


            ' Display the final password.
            lblPassword.Text = strFinal
            End If
    End Sub


    Private Sub txtWords_Enter(sender As Object, e As EventArgs) Handles txtWords.Enter
        txtWords.SelectAll()
    End Sub

    Private Sub txtWords_TextChanged(sender As Object, e As EventArgs) Handles txtWords.TextChanged
        lblPassword.Text = String.Empty
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub




End Class
