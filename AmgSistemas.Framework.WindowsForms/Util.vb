Imports System.Reflection
Imports System.Windows.Forms

Public Class Util

    Public Shared Function ConverterItems(ByVal Objetos As Object, ByVal Chave As String, ByVal Valor As String) As List(Of Item)

        Dim Items As List(Of Item) = Nothing

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then

            Items = New List(Of Item)

            For Each Item In Objetos
                Items.Add(New Item With {
                          .Identificador = RetornarValor(Item, Chave),
                          .Descricao = RetornarValor(Item, Valor)})
            Next

        End If

        Return Items
    End Function

    Public Shared Function PreencherCombobox(ByVal Combo As ComboBox,
                                             ByVal Objetos As List(Of Item)) As System.Windows.Forms.ComboBox

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then


            Combo.ValueMember = "Identificador"
            Combo.DisplayMember = "Descricao"

            For Each objeto In Objetos
                Combo.Items.Add(objeto)
            Next

        End If

        Return Combo
    End Function

    Public Shared Function PreencherCheckedListBox(ByVal objCheckedListBox As CheckedListBox,
                                             ByVal Objetos As List(Of Item)) As System.Windows.Forms.CheckedListBox

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then


            objCheckedListBox.ValueMember = "Identificador"
            objCheckedListBox.DisplayMember = "Descricao"

            For Each objeto In Objetos
                objCheckedListBox.Items.Add(objeto)
            Next

        End If

        Return objCheckedListBox
    End Function


    Public Shared Function PreencherDataGridCombobox(ByVal Combo As DataGridViewComboBoxColumn,
                                                     ByVal Objetos As List(Of Item)) As System.Windows.Forms.DataGridViewComboBoxColumn

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then


            Combo.ValueMember = "Identificador"
            Combo.DisplayMember = "Descricao"

            For Each objeto In Objetos
                Combo.Items.Add(objeto)
            Next

        End If

        Return Combo
    End Function

    Public Shared Function PreencherListBox(ByVal Listbox As ListBox,
                                             ByVal Objetos As List(Of Item)) As System.Windows.Forms.ListBox

        If Objetos IsNot Nothing AndAlso Objetos.Count > 0 Then

            Listbox.ValueMember = "Identificador"
            Listbox.DisplayMember = "Descricao"

            For Each objeto In Objetos
                Listbox.Items.Add(objeto)
            Next

        End If

        Return Listbox
    End Function

    Public Shared Function RetornarValor(ByVal Item As Object, ByVal NomePropriedade As String) As Object


        Dim propriedades_info As PropertyInfo() = GetType(Item).GetProperties

        For Each i In Item.GetType().GetProperties

            If i.Name = NomePropriedade Then
                Return i.GetValue(Item, i.GetIndexParameters)
            End If


        Next i

        Return Nothing
    End Function

    Public Shared Sub PreencherValor(ByRef Item As Object, ByVal NomePropriedade As String, ByVal Valor As String)


        Dim propriedades_info As PropertyInfo() = GetType(Item).GetProperties

        For Each i In Item.GetType().GetProperties

            If i.Name = NomePropriedade Then
                If i.GetType Is GetType(Boolean) Then
                    i.SetValue(Item, Convert.ToBoolean(Valor), i.GetIndexParameters)
                ElseIf i.GetType Is GetType(String) Then
                    i.SetValue(Item, Valor, i.GetIndexParameters)
                ElseIf i.GetType Is GetType(Int32) Then
                    i.SetValue(Item, Convert.ToInt32(Valor), i.GetIndexParameters)
                End If

                Exit For
            End If


        Next i

    End Sub

    Public Shared Function RecuperarValorComponente(ByVal Controle As System.Windows.Forms.Control, ByVal NomePropriedade As String) As String

        Dim Valor As String = String.Empty

        If Controle.GetType.Equals(GetType(System.Windows.Forms.ComboBox)) Then

            Dim Combo As System.Windows.Forms.ComboBox = DirectCast(Controle, ComboBox)

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

    Public Shared Function RecuperarItemSelecionado(ByVal Controle As Control, ByVal Objetos As Object, ByVal NomePropriedade As String) As Object

        Dim Item As Object = Nothing

        If Controle.GetType.Equals(GetType(System.Windows.Forms.ComboBox)) Then

            Dim Combo As System.Windows.Forms.ComboBox = DirectCast(Controle, ComboBox)

            If Combo.SelectedItem IsNot Nothing Then

                For Each objItem In Objetos

                    If RetornarValor(objItem, NomePropriedade) = Combo.SelectedItem.Identificador Then
                        Item = objItem
                    End If

                Next

            End If

        ElseIf Controle.GetType.Equals(GetType(System.Windows.Forms.CheckedListBox)) Then

            Dim objCheckedListBox As System.Windows.Forms.CheckedListBox = DirectCast(Controle, CheckedListBox)

            If objCheckedListBox.SelectedItem IsNot Nothing Then

                For Each objItem In Objetos

                    If RetornarValor(objItem, NomePropriedade) = objCheckedListBox.SelectedItem.Identificador Then
                        Item = objItem
                    End If

                Next

            End If
        Else
            Dim ListBox As System.Windows.Forms.ListBox = DirectCast(Controle, ListBox)

            If ListBox.SelectedItem IsNot Nothing Then

                For Each objItem In Objetos

                    If RetornarValor(objItem, NomePropriedade) = ListBox.SelectedItem.Identificador Then
                        Item = objItem
                    End If

                Next

            End If

        End If

        Return Item
    End Function

    Public Shared Function SelecionarItemControle(ByVal Controle As Control, ByVal Valor As String, ByVal NomePropriedade As String, Optional ByVal Valores As Object = Nothing) As Control

        Try

            If Controle.GetType.Equals(GetType(System.Windows.Forms.ComboBox)) Then

                Dim Combo As System.Windows.Forms.ComboBox = DirectCast(Controle, ComboBox)
                Dim Item As Object = (From objI In Combo.Items Where RetornarValor(objI, NomePropriedade) = Valor).FirstOrDefault()

                If Item IsNot Nothing Then
                    Combo.SelectedItem = Item
                    Return Combo
                End If

            ElseIf Controle.GetType.Equals(GetType(System.Windows.Forms.CheckedListBox)) Then

                Dim objCheckedListBox As CheckedListBox = DirectCast(Controle, CheckedListBox)

                If Valores Is Nothing OrElse Valores.Count = 0 Then

                    Dim Item As Object = (From objI In objCheckedListBox.Items Where RetornarValor(objI, NomePropriedade) = Valor).FirstOrDefault()

                    If Item IsNot Nothing Then
                        objCheckedListBox.SelectedItem = Item
                        Return objCheckedListBox
                    End If

                Else

                    If Valores IsNot Nothing AndAlso Valores.Count > 0 Then

                        Dim IndexesSelecionados As List(Of Integer) = New List(Of Integer)

                        For Each vlr In Valores

                            Dim index As Integer = 0

                            For Each Item In objCheckedListBox.Items

                                If RetornarValor(Item, NomePropriedade) = RetornarValor(vlr, NomePropriedade) Then
                                    IndexesSelecionados.Add(index)
                                End If

                                index += 1
                            Next

                        Next

                        For Each i In IndexesSelecionados
                            objCheckedListBox.SetItemChecked(i, True)
                        Next
                        Return objCheckedListBox
                    End If

                End If
            Else

                Dim objList As ListBox = DirectCast(Controle, ListBox)

                If objList.SelectionMode = SelectionMode.One Then

                    Dim Item As Object = (From objI In objList.Items Where RetornarValor(objI, NomePropriedade) = Valor).FirstOrDefault()

                    If Item IsNot Nothing Then
                        objList.SelectedItem = Item
                        Return objList
                    End If

                Else

                    If Valores IsNot Nothing AndAlso Valores.Count > 0 Then

                        For Each vlr In Valores

                            Dim Item As Object = (From objI In objList.Items Where RetornarValor(objI, NomePropriedade) = RetornarValor(vlr, NomePropriedade)).FirstOrDefault()

                            If Item IsNot Nothing Then
                                objList.SelectedItems.Add(Item)
                            End If

                        Next

                        Return objList
                    End If

                End If


            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return Controle
    End Function

    Public Shared Function RecuperarItemsSelecionados(Of Tipo)(ByVal Controle As Control, ByVal Objetos As Object, ByVal NomePropriedade As String) As List(Of Tipo)

        Dim Items As List(Of Tipo)

        If Controle.GetType.Equals(GetType(System.Windows.Forms.ListBox)) Then

            Dim ListBox As System.Windows.Forms.ListBox = DirectCast(Controle, ListBox)

            If ListBox.SelectedItems IsNot Nothing AndAlso ListBox.SelectedItems.Count > 0 Then

                For Each objItem In ListBox.SelectedItems

                    For Each objItemObjeto In Objetos

                        If RetornarValor(objItemObjeto, NomePropriedade) = objItem.Identificador Then

                            If Items Is Nothing Then Items = New List(Of Tipo)

                            Items.Add(objItemObjeto)

                        End If

                    Next

                Next

            End If

        ElseIf Controle.GetType.Equals(GetType(System.Windows.Forms.CheckedListBox)) Then

            Dim ListBox As System.Windows.Forms.CheckedListBox = DirectCast(Controle, CheckedListBox)

            If ListBox.CheckedItems IsNot Nothing AndAlso ListBox.CheckedItems.Count > 0 Then

                For Each objItem In ListBox.CheckedItems

                    For Each objItemObjeto In Objetos

                        If RetornarValor(objItemObjeto, NomePropriedade) = objItem.Identificador Then

                            If Items Is Nothing Then Items = New List(Of Tipo)

                            Items.Add(objItemObjeto)

                        End If

                    Next

                Next

            End If
        End If

        Return Items
    End Function

End Class

Public Class Item

    Public Property Identificador As String
    Public Property Descricao As String

End Class
