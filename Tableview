Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Drawing
Imports System

Public Class TableView
    Inherits ListView

    Public Property AlternateBackColor As Color
    Public Property AlternameFontColor As Color

    ''' <summary>
    ''' TableView Constructor Method.
    ''' Set some default properties for the TableView Object.
    ''' </summary>
    Sub New()
        Font = New Font("Courier New", 8)
        View = Windows.Forms.View.Details
        GridLines = True
        FullRowSelect = True
        MultiSelect = True
        HideSelection = True

        'Prevents screen flicker
        Me.DoubleBuffered = True

        AddHandler Me.ColumnWidthChanging, AddressOf OnTableViewColumnWidthChanging
        AddHandler Me.ColumnWidthChanged, AddressOf OnTableViewColumnWidthChanged

    End Sub

    ''' <summary>
    ''' Creates Column headers using an array of strings.
    ''' </summary>
    ''' <param name="StringLabels">String array</param>
    Public Sub CreateColumnHeaders(StringLabels As String())
        Dim ColumnWidth As Integer = Me.Size.Width / StringLabels.Count
        Dim index As Integer = 0
        For Each LabelString As String In StringLabels
            If index <> StringLabels.Count - 1 Then
                Me.Columns.Add(LabelString, ColumnWidth, HorizontalAlignment.Left)
            Else
                Me.Columns.Add(LabelString, -2, HorizontalAlignment.Left)
            End If
            index = index + 1
        Next
    End Sub

    ''' <summary>
    ''' Creates a row using an array of strings.
    ''' Each array slot will build a cell.
    ''' </summary>
    ''' <param name="StringLabels">String array</param>
    Public Sub CreateRow(StringLabels As String())
        Dim booleanFirstColumn As Boolean = True
        Dim booleanToggle As Boolean = True
        Dim row As New ListViewItem(StringLabels(0), 0)

        'If the labels array does not equal the same size as the amount of column headers.
        'Then add blank labels to the array to match the size of the column headers.
        If StringLabels.Count <> Me.Columns.Count Then
            For index As Integer = 0 To Me.Columns.Count - 1
                If index > StringLabels.Count - 1 Then
                    ReDim Preserve StringLabels(index)
                    StringLabels(index) = ""
                End If
            Next
        End If

        'for each label create a cell
        For Each LabelString As String In StringLabels
            If booleanFirstColumn Then
                booleanFirstColumn = False
            Else
                row.SubItems.Add(LabelString)
            End If
        Next
        Me.Items.AddRange({row})
    End Sub

    ''' <summary>
    ''' Will highlight a column using the column index.
    ''' </summary>
    ''' <param name="ColumnColor">System.Drawing.Color</param>
    ''' <param name="ColumnIndex">Integer</param>
    Public Sub HighlightColumn(ColumnColor As Color, fontColor As Color, ColumnIndex As Integer)
        For Each row As ListViewItem In Me.Items
            row.UseItemStyleForSubItems = False
            row.SubItems(ColumnIndex).BackColor = ColumnColor
            row.SubItems(ColumnIndex).ForeColor = fontColor
        Next
    End Sub

    ''' <summary>
    ''' Will highlight a row using the row index.
    ''' </summary>
    ''' <param name="RowColor">System.Drawing.Color</param>
    ''' <param name="RowIndex">Integer</param>
    Public Sub HighlightRow(RowColor As Color, fontColor As Color, RowIndex As Integer)
        Me.Items(RowIndex).UseItemStyleForSubItems = False
        For Each column As ListViewItem.ListViewSubItem In Me.Items(RowIndex).SubItems
            column.BackColor = RowColor
            column.ForeColor = fontColor
        Next
    End Sub

    ''' <summary>
    ''' Will highlight a cell at the given row and column index.
    ''' </summary>
    ''' <param name="CellColor">System.Drawing.Color</param>
    ''' <param name="Rowindex">Integer</param>
    ''' <param name="ColumnIndex">Integer</param>
    Public Sub HighlightCell(CellColor As Color, fontColor As Color, Rowindex As Integer, ColumnIndex As Integer)
        Dim row As ListViewItem = Me.Items(Rowindex)
        row.UseItemStyleForSubItems = False
        row.SubItems(ColumnIndex).BackColor = CellColor
        row.SubItems(ColumnIndex).ForeColor = fontColor
    End Sub

    ''' <summary>
    ''' Will highlight rows in an alternating pattern.
    ''' </summary>
    Public Sub HighlightAlernatingRows()
        Dim booleanAlternate As Boolean = True
        Dim index As Integer = 0
        For Each row As ListViewItem In Me.Items
            If booleanAlternate Then
                Me.HighlightRow(Me.BackColor, Me.ForeColor, index)
                booleanAlternate = False
            Else
                Me.HighlightRow(AlternateBackColor, AlternameFontColor, index)
                booleanAlternate = True
            End If
            index = index + 1
        Next
    End Sub

    ''' <summary>
    ''' Will highlight columns in an alternating pattern.
    ''' </summary>
    Public Sub HighlightAlernatingColumns()
        Dim booleanAlternate As Boolean = True
        Dim index As Integer = 0
        For Each row As ListViewItem.ListViewSubItem In Me.Items(0).SubItems
            If booleanAlternate Then
                Me.HighlightColumn(Me.BackColor, Me.ForeColor, index)
                booleanAlternate = False
            Else
                Me.HighlightColumn(AlternateBackColor, AlternameFontColor, index)
                booleanAlternate = True
            End If
            index = index + 1
        Next
    End Sub

    Public Sub OnTableViewColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs)
        'Dim size As Integer = 0
        'For Each column As ColumnHeader In Me.Columns
        'Size = Size + column.Width
        'Next
        'If e.ColumnIndex <> Me.Columns.Count - 1 Then
        'Me.Columns.Item(Me.Columns.Count - 1).Width = -2
        'End If

        If e.NewWidth < 50 Then
            e.NewWidth = 50
            e.Cancel = True
        End If

    End Sub

    Public Sub OnTableViewColumnWidthChanged(send As Object, e As ColumnWidthChangedEventArgs)
        'If e.ColumnIndex <> Me.Columns.Count - 1 Then
        ' Me.Columns.Item(Me.Columns.Count - 1).Width = -2
        'End If
        If Me.Columns.Item(e.ColumnIndex).Width < 50 Then
            Me.Columns.Item(e.ColumnIndex).Width = 50
        End If
    End Sub

End Class

