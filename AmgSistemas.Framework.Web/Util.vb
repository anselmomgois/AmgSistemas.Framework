Imports System.Web.UI.WebControls


Public Class Util

    Public Shared Function ConverterItems(ByVal Objetos As Object, ByVal Chave As String, ByVal Valor As String) As List(Of ListItem)

        Dim Items As List(Of ListItem) = Nothing

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then

            Items = New List(Of ListItem)

            For Each Item In Objetos
                Items.Add(New ListItem With { _
                          .Value = RetornarValor(Item, Chave), _
                          .Text = RetornarValor(Item, Valor)})
            Next

        End If

        Return Items
    End Function

    Public Shared Function PreencherDropDown(ByVal DropDown As DropDownList,
                                             ByVal Objetos As List(Of ListItem), _
                                             ByVal AdicionarItemSelecionar As Boolean) As DropDownList

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then

            DropDown.Items.Clear()

            If AdicionarItemSelecionar Then DropDown.Items.Add("Selecione")

            For Each objeto In Objetos
                DropDown.Items.Add(objeto)
            Next

        End If

        Return DropDown
    End Function

    Public Shared Function PreencherListBox(ByVal Listbox As ListBox,
                                             ByVal Objetos As List(Of ListItem)) As ListBox

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then

            For Each objeto In Objetos
                Listbox.Items.Add(objeto)
            Next

        End If

        Return Listbox
    End Function

    Public Shared Function RecuperarValorComponente(ByVal Controle As System.Web.UI.Control, ByVal NomePropriedade As String) As String

        Dim Valor As String = String.Empty

        If Controle.GetType.Equals(GetType(DropDownList)) Then

            Dim Combo As DropDownList = DirectCast(Controle, DropDownList)

            If Combo.SelectedItem IsNot Nothing Then
                Valor = RetornarValor(Combo.SelectedItem, NomePropriedade)
            End If

        Else

            Dim ListBox As ListBox = DirectCast(Controle, ListBox)

            If ListBox.SelectedItem IsNot Nothing Then
                Valor = RetornarValor(ListBox.SelectedItem, NomePropriedade)
            End If

        End If

        Return Valor
    End Function

    Public Shared Function RecuperarItemSelecionado(ByVal Controle As System.Web.UI.Control, ByVal Objetos As Object, ByVal NomePropriedade As String) As Object

        Dim Item As Object = Nothing

        If Controle.GetType.Equals(GetType(DropDownList)) Then

            Dim Combo As DropDownList = DirectCast(Controle, DropDownList)

            If Combo.SelectedItem IsNot Nothing Then

                For Each objItem In Objetos

                    If RetornarValor(objItem, NomePropriedade) = Combo.SelectedItem.Value Then
                        Item = objItem
                    End If

                Next

            End If

        Else
            Dim ListBox As ListBox = DirectCast(Controle, ListBox)

            If ListBox.SelectedItem IsNot Nothing Then

                For Each objItem In Objetos

                    If RetornarValor(objItem, NomePropriedade) = ListBox.SelectedItem.Value Then
                        Item = objItem
                    End If

                Next

            End If

        End If

        Return Item
    End Function

    Public Shared Function SelecionarItemControle(ByVal Controle As System.Web.UI.Control, ByVal Valor As String, ByVal NomePropriedade As String, Optional ByVal Valores As Object = Nothing) As System.Web.UI.Control

        Try

        If Controle.GetType.Equals(GetType(DropDownList)) Then

            Dim Combo As DropDownList = DirectCast(Controle, DropDownList)
            Dim Item As ListItem = (From objI As ListItem In Combo.Items Where objI.Value = Valor).FirstOrDefault()

            If Item IsNot Nothing Then
                Combo.SelectedValue = Item.Value
                Return Combo
            End If

        Else

            Dim objList As ListBox = DirectCast(Controle, ListBox)

            If objList.SelectionMode = ListSelectionMode.Single Then

                Dim Item As ListItem = (From objI As ListItem In objList.Items Where objI.Value = Valor).FirstOrDefault()

                If Item IsNot Nothing Then
                    objList.SelectedValue = Item.Value
                    Return objList
                End If

            Else

                If Valores IsNot Nothing AndAlso Valores.Count > 0 Then

                    For Each vlr In Valores

                        Dim Item As ListItem = (From objI As ListItem In objList.Items Where objI.Value = RetornarValor(vlr, NomePropriedade)).FirstOrDefault()

                        If Item IsNot Nothing Then
                            Item.Selected = True
                        End If

                    Next

                    Return objList
                End If

            End If


        End If

        Catch ex As Exception

        End Try

        Return Controle
    End Function

    Public Shared Function RecuperarItemsSelecionados(Of Tipo)(ByVal Controle As System.Web.UI.Control, ByVal Objetos As Object, ByVal NomePropriedade As String) As List(Of Tipo)

        Dim Items As List(Of Tipo)

        If Controle.GetType.Equals(GetType(ListBox)) Then

            Dim ListBox As ListBox = DirectCast(Controle, ListBox)

            If ListBox.Items IsNot Nothing AndAlso ListBox.Items.Count > 0 AndAlso (From item As ListItem In ListBox.Items
                                                                                    Where item.Selected).Count > 0 Then

                For Each objItem In (From item As ListItem In ListBox.Items Where item.Selected)

                    For Each objItemObjeto In Objetos

                        If RetornarValor(objItemObjeto, NomePropriedade) = objItem.Value Then

                            If Items Is Nothing Then Items = New List(Of Tipo)

                            Items.Add(objItemObjeto)

                        End If

                    Next

                Next

            End If

        End If

        Return Items
    End Function

    Public Shared Function RetornarValor(ByVal Item As Object, ByVal NomePropriedade As String) As Object

        Select Case NomePropriedade.ToUpper

            Case "IDENTIFICADOR"
                Return Item.Identificador
            Case "CODIGO"
                Return Item.Codigo
            Case "DESCRICAO"
                Return Item.Descricao
            Case "NOME"
                Return Item.Nome
            Case "NOMECODIGO"
                Return Item.NomeCodigo
            Case "DESNOME"
                Return Item.DesNome
        End Select

        Return Nothing
    End Function

End Class
