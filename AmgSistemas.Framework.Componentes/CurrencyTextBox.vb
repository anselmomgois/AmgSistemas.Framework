Option Explicit On

Imports System.ComponentModel
Imports System.Globalization
Imports System.Windows.Forms

Public Class CurrencyTextBox

#Region "Variables"

    Protected blnSelectTxt As Boolean = True
    Private selLength As Integer
    Private intFirstKey As Integer = 0
    Private intDecimal As Integer = NumberFormatInfo.CurrentInfo.CurrencyDecimalDigits
    Private strDecimal As String = ""
    Private strCurrencySymbol As String = String.Empty 'NumberFormatInfo.CurrentInfo.CurrencySymbol
    Private strDecimalSeparator As String = NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator
    Private strGroupSeparator As String = NumberFormatInfo.CurrentInfo.CurrencyGroupSeparator
    Private _txt As System.Windows.Forms.TextBox

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByRef txt As System.Windows.Forms.TextBox)

        Me.Inicializar(txt)

    End Sub

#End Region

#Region "Events"

    Public Sub removeevent()
        If _txt Is Nothing Then
            Exit Sub
        End If
        RemoveHandler _txt.GotFocus, AddressOf OnGotFocus
        RemoveHandler _txt.KeyDown, AddressOf OnKeyDown
        RemoveHandler _txt.KeyPress, AddressOf OnKeyPress
        RemoveHandler _txt.Enter, AddressOf OnEnter
    End Sub

    Protected Sub OnGotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.SelTxtOnEnter = True Then _txt.SelectAll()
    End Sub

    Protected Sub OnKeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        Dim intSelPos As Integer = _txt.SelectionStart
        selLength = _txt.SelectionLength
        Dim s As String = _txt.Text
        Dim p As Integer = 0
        Dim intrelSelPos As Integer
        Dim i As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Delete Then
            intrelSelPos = RelativePos(s, intSelPos)
            If intSelPos = _txt.TextLength - Me.DecimalsDigits - 1 Then intSelPos += 1
            s = ExtractNumbers(s)
            If selLength = _txt.TextLength Then
                _txt.Text = Me.CurrencySymbol & "0" & strDecimalSeparator & strDecimal
                _txt.SelectionStart = Me.CurrencySymbol.Length + 1
                _txt.SelectionLength = 0
                e.Handled = True
                Exit Sub
            Else
                If selLength > 0 Then
                    s = s.Remove(intrelSelPos, Len(ExtractNumbers(_txt.SelectedText)))
                    If _txt.SelectionStart + selLength > _txt.TextLength - DecimalsDigits Then
                        For i = 1 To (_txt.SelectionStart + selLength) - (_txt.TextLength - DecimalsDigits)
                            s = s & "0"
                        Next
                    End If
                Else
                    If _txt.SelectionStart >= _txt.TextLength - DecimalsDigits - 1 Then                        
                        s = s.Remove(intrelSelPos, 1) & "0"
                    Else
                        s = s.Remove(intrelSelPos, 1)
                    End If
                End If

                If _txt.SelectionStart > _txt.TextLength - DecimalsDigits Then
                    _txt.Text = PutMask(s)
                    intSelPos -= 1
                ElseIf _txt.SelectionStart = _txt.TextLength - DecimalsDigits Then
                    intSelPos -= 1
                    If s.Length = Me.DecimalsDigits Then
                        s = "0" & s
                        intSelPos += 1
                    End If
                    _txt.Text = PutMask(s)
                Else
                    If s.Length = Me.DecimalsDigits Then
                        s = "0" & s
                        intSelPos += 1
                    End If
                    intSelPos -= 1
                    _txt.Text = PutMask(s)
                End If
                If Len(s) > Me.DecimalsDigits Then
                    If selLength > 0 Then intSelPos += 1
                    _txt.SelectionStart = intSelPos '+ 1
                Else
                    _txt.SelectionStart = _txt.TextLength - DecimalsDigits
                    Beep()
                End If
                e.Handled = True
                Exit Sub
            End If
        End If

    End Sub

    Protected Sub OnKeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim intSelPos As Integer = _txt.SelectionStart
        Dim selLength As Integer = _txt.SelectionLength
        Dim c As String
        Dim s As String = _txt.Text
        Dim i As Integer
        Dim ultimaPosicao As Boolean = False

        On Error Resume Next

        Dim intrelSelPos As Integer
        Dim blnLeft As Boolean = False
        c = ""
        intrelSelPos = RelativePos(s, intSelPos)
        s = ExtractNumbers(s)
        If intFirstKey = 0 And s = strDecimal & "0" Then
            intFirstKey += 1
        End If

        If Asc(e.KeyChar) = Keys.Back Then
            Dim intInitialLength As Integer = _txt.TextLength
            Dim intInitialPos As Integer = intSelPos
            If selLength = _txt.TextLength Then
                _txt.Text = Me.CurrencySymbol & "0" & strDecimalSeparator & strDecimal
                _txt.SelectionStart = Me.CurrencySymbol.Length + 1
                _txt.SelectionLength = 0
                e.Handled = True
                Exit Sub
            Else
                If selLength > 0 Then
                    s = s.Remove(intrelSelPos, Len(ExtractNumbers(_txt.SelectedText)))
                    If _txt.SelectionStart + selLength > _txt.TextLength - DecimalsDigits Then
                        For i = 1 To (_txt.SelectionStart + selLength) - (_txt.TextLength - DecimalsDigits)
                            s = s & "0"
                        Next
                    End If
                Else
                    If intrelSelPos > 0 Then s = s.Remove(intrelSelPos - 1, 1)
                End If
                If _txt.SelectionStart > _txt.TextLength - DecimalsDigits Then
                    s = s & "0"
                    _txt.Text = PutMask(s)
                    intSelPos -= 1
                ElseIf _txt.SelectionStart = _txt.TextLength - DecimalsDigits Then
                    intSelPos -= 2
                    If s.Length = Me.DecimalsDigits Then
                        s = "0" & s
                        intSelPos += 1
                    End If
                    _txt.Text = PutMask(s)
                Else
                    blnLeft = True
                    If s.Length = Me.DecimalsDigits Then
                        s = "0" & s
                        intSelPos += 1
                    End If
                    intSelPos -= 1
                    _txt.Text = PutMask(s)
                End If
                If Len(s) > Me.DecimalsDigits Then
                    If selLength > 0 Then intSelPos += 1
                    If blnLeft Then
                        _txt.SelectionStart = intInitialPos - (intInitialLength - _txt.TextLength)
                    Else
                        _txt.SelectionStart = intSelPos
                    End If

                Else
                    _txt.SelectionStart = _txt.TextLength - DecimalsDigits - 1
                    Beep()
                End If
                e.Handled = True
                Exit Sub
            End If
        End If
        If Asc(e.KeyChar) < 33 Then
            e.Handled = True
            Exit Sub
        End If
        If Asc(e.KeyChar) = 44 Or Asc(e.KeyChar) = 46 Then 'Igual a . ou ,
            _txt.DeselectAll()
            If _txt.SelectionStart > _txt.TextLength - DecimalsDigits - 1 Then
                _txt.SelectionStart = _txt.TextLength - DecimalsDigits - 1
            Else
                _txt.SelectionStart = _txt.TextLength - DecimalsDigits
            End If
        End If
        If Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57 Then
            c = Chr(Asc(e.KeyChar))
            If intFirstKey = 1 And c <> "0" Then
                'intFirstKey += 1
                'intrelSelPos = 0
                's = strDecimal
            End If
            If intSelPos = 0 And s.Length = 0 Then
                s = c
            ElseIf Len(s) >= 14 Then
                ' original-----------------
                'If intrelSelPos = s.Length Then intrelSelPos -= 1
                's = s.Remove(intrelSelPos, 1)
                's = s.Insert(intrelSelPos, c)
                '-----------------
                If Len(s) = intrelSelPos Then
                    ultimaPosicao = True
                    s = s.Remove(intrelSelPos, 1)
                Else
                    s = s.Remove(intrelSelPos, 1)
                    s = s.Insert(intrelSelPos, c)

                    If _txt.TextLength = _txt.SelectionLength Then
                        s = c
                    End If
                End If
            ElseIf selLength > 0 Then
                If _txt.SelectionStart > _txt.TextLength - DecimalsDigits - 1 Then
                    s = s.Remove(intrelSelPos, selLength)
                    If selLength = 1 Then
                        s = s.Insert(intrelSelPos, c)
                    ElseIf selLength > 0 Then
                        s = s.Insert(intrelSelPos, c)
                        s = s & strDecimal.Substring(1, selLength - 1)
                    End If
                ElseIf _txt.SelectionStart + selLength > _txt.TextLength - DecimalsDigits Then
                    If _txt.SelectionLength = _txt.TextLength Then
                        s = c
                    Else
                        s = s.Remove(intrelSelPos, Len(ExtractNumbers(_txt.SelectedText)))
                        s = s.Insert(intrelSelPos, c)
                        If s.Length < Me.DecimalsDigits + 1 Then
                            For i = 1 To (_txt.SelectionStart + selLength) - (_txt.TextLength - DecimalsDigits)
                                s = s & "0"
                            Next
                        End If
                    End If
                Else
                    s = s.Remove(intrelSelPos, Len(ExtractNumbers(_txt.SelectedText)))
                    s = s.Insert(intrelSelPos, c)
                End If
            ElseIf _txt.SelectionStart > _txt.TextLength - DecimalsDigits - 1 Then
                ' original-----------------
                'If Len(s) = intrelSelPos Then intrelSelPos -= DecimalsDigits
                's = s.Remove(intrelSelPos, 1)
                's = s.Insert(intrelSelPos, c)
                ' -----------------
                If Len(s) = intrelSelPos Then
                    ultimaPosicao = True
                        s = s.Remove(intrelSelPos, 1)
                Else
                    s = s.Remove(intrelSelPos, 1)
                    s = s.Insert(intrelSelPos, c)
                End If
            Else
                s = s.Insert(intrelSelPos, c)
                If s.Length = Me.DecimalsDigits + 2 And s.Substring(0, 1) = "0" Then
                    intrelSelPos = 0
                End If
            End If
        End If
        e.Handled = True
        _txt.Text = PutMask(s)

        ' caso não seja a ultima posição
        If Not ultimaPosicao Then
            ' obter a nova posição
            intrelSelPos = RelativePos2(_txt.Text, intrelSelPos)
        Else
            ' setar ultima posição de forma que o usuário não consiga digitar outro valor.
            intrelSelPos = Len(_txt.Text)
        End If

        If c = "" Then
            If Asc(e.KeyChar) <> 44 And Asc(e.KeyChar) <> 46 Then Beep()
        Else
            If intSelPos = 0 Then
                _txt.SelectionStart = Me.CurrencySymbol.Length + 1
            Else
                _txt.SelectionStart = intrelSelPos
            End If
        End If

    End Sub

    Protected Sub OnEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        intFirstKey = 0
    End Sub

    Protected Sub OnCreateControl()
        SetCurrency()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property Author() As String
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property Version() As String
        Get
            With System.Reflection.Assembly.GetAssembly(Me.GetType()).GetName().Version
                Return .Major.ToString & "." & .Minor.ToString & "." _
                     & .Build.ToString & "." & .Revision().ToString()
            End With
        End Get
    End Property

    <CategoryAttribute("Currency"), _
       DescriptionAttribute("Number of digits to the right of the decimal point.")> _
   Public Property DecimalsDigits() As Integer
        Get
            Return intDecimal
        End Get
        Set(ByVal Value As Integer)
            intDecimal = Value
            SetCurrency()
        End Set
    End Property

    <CategoryAttribute("Currency"), _
    DescriptionAttribute("Decimal Separator")> _
    Public Property CurrencyDecimalSeparator() As String
        Get
            Return strDecimalSeparator
        End Get

        Set(ByVal Value As String)
            If Value = "," Or Value = "." Then
                strDecimalSeparator = Value
                SetCurrency()
            End If
        End Set
    End Property

    <CategoryAttribute("Currency"), _
    DescriptionAttribute("Group Separator")> _
    Public Property CurrencyGroupSeparator() As String
        Get
            Return strGroupSeparator
        End Get

        Set(ByVal Value As String)
            If Value = "," Or Value = "." Then
                strGroupSeparator = Value
                SetCurrency()
            End If
        End Set
    End Property

    <CategoryAttribute("Currency"), _
    DescriptionAttribute("Currency Simbol")> _
    Public Property CurrencySymbol() As String
        Get
            Return strCurrencySymbol
        End Get

        Set(ByVal Value As String)
            strCurrencySymbol = Value
            SetCurrency()
        End Set
    End Property

    <CategoryAttribute("Currency"), _
       DefaultValueAttribute(True), _
       DescriptionAttribute("Select All The text On Focus")> _
       Public Property SelTxtOnEnter() As Boolean
        Get
            Return blnSelectTxt
        End Get
        Set(ByVal Value As Boolean)
            blnSelectTxt = Value
        End Set
    End Property

#End Region

#Region "Subs"

    Private Sub SetCurrency()
        Dim i As Integer
        strDecimal = ""
        For i = 1 To Me.DecimalsDigits
            strDecimal = "0" & strDecimal
        Next
        _txt.Text = Me.CurrencySymbol & "0" & strDecimalSeparator & strDecimal
    End Sub

#End Region

#Region "Functions"

    Public Sub Inicializar(ByRef txt As TextBox)

        _txt = txt
        removeevent()

        If txt.ReadOnly OrElse Not txt.Enabled Then
            Exit Sub
        End If

        AddHandler _txt.GotFocus, AddressOf OnGotFocus
        AddHandler _txt.KeyDown, AddressOf OnKeyDown
        AddHandler _txt.KeyPress, AddressOf OnKeyPress
        AddHandler _txt.Enter, AddressOf OnEnter

        Dim aux As String = _txt.Text
        OnCreateControl()
        If aux <> String.Empty Then
            _txt.Text = aux
        End If

    End Sub

    Protected Function ExtractNumbers(ByVal t As String) As String
        Dim i As Integer
        Dim r As String = ""
        For i = 0 To t.Length - 1
            If Asc(t.Substring(i, 1)) >= 48 And Asc(t.Substring(i, 1)) <= 57 Then r = r & t.Substring(i, 1)
        Next
        Return r
    End Function

    Protected Function RelativePos(ByVal t As String, ByVal p As Integer) As Integer
        Dim i As Integer
        For i = 0 To p - 1
            If Asc(t.Substring(i, 1)) >= 48 And Asc(t.Substring(i, 1)) <= 57 Then
                RelativePos += 1
            End If
        Next
    End Function

    Protected Function RelativePos2(ByVal t As String, ByVal p As Integer) As Integer
        Dim i As Integer
        Dim pp As Integer
        Dim ppp As Integer
        For i = 0 To t.Length - 1
            pp += 1
            If Asc(t.Substring(i, 1)) >= 48 And Asc(t.Substring(i, 1)) <= 57 Then
                ppp += 1
            End If
            If ppp > p Then
                If p = 0 Then
                    RelativePos2 = t.Length - DecimalsDigits - 1
                Else
                    RelativePos2 = pp
                End If
                Exit For
            End If
        Next
    End Function

    Protected Function PutMask(ByVal t As String) As String
        Dim r As String = ""

        If t.Length < Me.DecimalsDigits + 1 Then t = t & strDecimal
        t = t.Insert(t.Length - Me.DecimalsDigits, NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator)
        r = Format(CDbl(t), Me.CurrencySymbol & "#,##0." & strDecimal)
        If Not Me.CurrencyGroupSeparator = NumberFormatInfo.CurrentInfo.CurrencyGroupSeparator Then
            r = r.Replace(NumberFormatInfo.CurrentInfo.CurrencyGroupSeparator, Me.CurrencyGroupSeparator)
        End If
        If Not r.Substring(r.Length - Me.DecimalsDigits - 1, 1) = Me.CurrencyDecimalSeparator Then
            r = r.Remove(r.Length - Me.DecimalsDigits - 1, 1)
            r = r.Insert(r.Length - Me.DecimalsDigits, Me.CurrencyDecimalSeparator)
        End If
        Return r
    End Function

    Public Function CurrencyValue() As String
        Return _txt.Text.Substring(Me.CurrencySymbol.Length + 1, _txt.TextLength - Me.CurrencySymbol.Length - 1)
    End Function

#End Region

End Class