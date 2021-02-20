Imports System.Windows.Forms

Public Class ControleColumnData
    Inherits DataGridViewColumn

    Public Sub New()
        MyBase.New(New ColumnDataEditCell())
    End Sub

#Region "[ATRIBUTOS]"

    Private _Format As System.Windows.Forms.DateTimePickerFormat = DateTimePickerFormat.Short
    Private _CustomFormat As String = String.Empty
    'Private _MinDate As System.DateTime = DateTimePicker.MinimumDateTime
    'Private _MaxDate As System.DateTime = DateTimePicker.MaximumDateTime

#End Region

#Region "[PROPRIEDADES]"

    Public Overrides Property CellTemplate() As DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As DataGridViewCell)

            ' Ensure that the cell used for the template is a CalendarCell.
            If Not (value Is Nothing) AndAlso _
                Not value.GetType().IsAssignableFrom(GetType(ColumnDataEditCell)) _
                Then
                Throw New InvalidCastException("Must be a CustonEditCell")
            End If
            MyBase.CellTemplate = value

        End Set
    End Property

    Public Property cFormat() As System.Windows.Forms.DateTimePickerFormat
        Get
            Return _Format
        End Get
        Set(ByVal value As System.Windows.Forms.DateTimePickerFormat)
            _Format = value
        End Set
    End Property

    Public Property CustomFormat() As String
        Get
            Return _CustomFormat
        End Get
        Set(ByVal value As String)
            _CustomFormat = value
        End Set
    End Property

    Private ReadOnly Property ColumnDataEditCellTemplate() As ColumnDataEditCell
        Get
            Return TryCast(Me.CellTemplate, ColumnDataEditCell)
        End Get
    End Property

    'Public Property cMinDate() As System.DateTime
    '    Get
    '        Return _MinDate
    '    End Get
    '    Set(ByVal value As System.DateTime)
    '        _MinDate = value
    '    End Set
    'End Property

    'Public Property cMaxDate() As System.DateTime
    '    Get
    '        Return _MaxDate
    '    End Get
    '    Set(ByVal value As System.DateTime)
    '        _MaxDate = value
    '    End Set
    'End Property

#End Region

End Class

Public Class ColumnDataEditCell
    Inherits DataGridViewTextBoxCell

    Public Sub New()
    End Sub

    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, _
                                                  ByVal initialFormattedValue As Object, _
                                                  ByVal dataGridViewCellStyle As DataGridViewCellStyle)



        ' Set the value of the editing control to the current cell value.
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)

        Dim ctl As ColumnDataEditingControl = CType(DataGridView.EditingControl, ColumnDataEditingControl)
        Dim mecol As New ControleColumnData '= Me.OwningColumn        

        mecol = TryCast(Me.OwningColumn, ControleColumnData)        

        With ctl

            .Format = DateTimePickerFormat.Custom

            If mecol IsNot Nothing Then
                .Format = mecol.cFormat
            End If

            If .Format = DateTimePickerFormat.Custom AndAlso mecol IsNot Nothing Then

                .CustomFormat = mecol.CustomFormat

            Else

                .CustomFormat = Me.Style.Format

            End If

            '.MinDate = mecol.cMinDate
            '.MaxDate = mecol.cMaxDate

        End With

    End Sub

    Public Overrides ReadOnly Property EditType() As Type
        Get
            Return GetType(ColumnDataEditingControl)
        End Get
    End Property

    Public Overrides ReadOnly Property ValueType() As Type
        Get
            ' Return the type of the value that CalendarCell contains.
            Return GetType(String)
        End Get
    End Property

    Public Overrides ReadOnly Property DefaultNewRowValue() As Object
        Get
            ' Use the current date and time as the default value.
            Return ""
        End Get
    End Property

    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal cellState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)
    End Sub

End Class

Class ColumnDataEditingControl
    Inherits DateTimePicker
    Implements IDataGridViewEditingControl

    Private dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean = False
    Private rowIndexNum As Integer

    Public Sub New()

    End Sub

    Public Property EditingControlFormattedValue() As Object _
        Implements IDataGridViewEditingControl.EditingControlFormattedValue

        Get
            Return Me.Text
        End Get

        Set(ByVal value As Object)
            Me.Text = value.ToString
        End Set

    End Property

    Public Function EditingControlWantsInputKey(ByVal key As Keys, _
           ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
           Implements IDataGridViewEditingControl.EditingControlWantsInputKey

        Return True
    End Function

    Public Function GetEditingControlFormattedValue(ByVal context _
        As DataGridViewDataErrorContexts) As Object _
        Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

        Return Me.Text
    End Function

    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As  _
        DataGridViewCellStyle) _
        Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

        Me.Font = dataGridViewCellStyle.Font
        Me.ForeColor = dataGridViewCellStyle.ForeColor
        Me.BackColor = dataGridViewCellStyle.BackColor

    End Sub

    Public Property EditingControlRowIndex() As Integer _
        Implements IDataGridViewEditingControl.EditingControlRowIndex

        Get
            Return rowIndexNum
        End Get
        Set(ByVal value As Integer)
            rowIndexNum = value
        End Set

    End Property

    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
        Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

        ' No preparation needs to be done.

    End Sub

    Public ReadOnly Property RepositionEditingControlOnValueChange() _
        As Boolean Implements _
        IDataGridViewEditingControl.RepositionEditingControlOnValueChange

        Get
            Return False
        End Get

    End Property

    Public Property EditingControlDataGridView() As DataGridView _
        Implements IDataGridViewEditingControl.EditingControlDataGridView

        Get
            Return dataGridViewControl
        End Get
        Set(ByVal value As DataGridView)
            dataGridViewControl = value
        End Set

    End Property

    Public Property EditingControlValueChanged() As Boolean _
        Implements IDataGridViewEditingControl.EditingControlValueChanged

        Get
            Return valueIsChanged
        End Get
        Set(ByVal value As Boolean)
            valueIsChanged = value
        End Set

    End Property

    Public ReadOnly Property EditingControlCursor() As Cursor _
        Implements IDataGridViewEditingControl.EditingPanelCursor

        Get
            Return MyBase.Cursor
        End Get

    End Property

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)

        ' Notify the DataGridView that the contents of the cell have changed.
        valueIsChanged = True
        Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
        MyBase.OnTextChanged(e)

    End Sub

    ''' <summary>
    ''' Evento que avisa que o valor está sendo editado, corrige bug que apagava data
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' [octavio.piramo] 12/09/2008 Criado
    ''' </history>
    Private Sub ColumnDataEditingControl_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        ' Notify the DataGridView that the contents of the cell have changed.
        valueIsChanged = True
        Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
        MyBase.OnTextChanged(e)

    End Sub

End Class