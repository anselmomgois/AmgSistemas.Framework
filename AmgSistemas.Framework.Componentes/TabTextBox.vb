Imports System
Imports System.Windows.Forms

Public Class TabTextBox
    Inherits System.Windows.Forms.TextBox

    ''' <summary>
    ''' Sobrescreve método IsInputKey do TextBox para que o evento
    ''' KeyDown ou o evento de KeyUp seja ativado ao pressionar TAB
    ''' </summary>
    ''' <param name="keyData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function IsInputKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean

        If keyData = Keys.Tab Then
            Return True
        Else
            Return MyBase.IsInputKey(keyData)
        End If

    End Function

End Class
