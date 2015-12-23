
Public Delegate Sub DelgReportViewer(ByVal ds As System.Data.DataSet, ByVal RptMain As CrystalDecisions.CrystalReports.Engine.ReportClass)
Public Class frmReportViewer
    Inherits System.Windows.Forms.Form
    Public oDBDataSource As SAPbouiCOM.DBDataSource
    Public oForm As SAPbouiCOM.Form
    Public oCompany As New SAPbobsCOM.Company
    Public iniViewer As DelgReportViewer


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents rptViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.rptViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'rptViewer
        '
        Me.rptViewer.ActiveViewIndex = -1
        Me.rptViewer.AutoScroll = True
        Me.rptViewer.DisplayGroupTree = False
        Me.rptViewer.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rptViewer.Location = New System.Drawing.Point(0, 0)
        Me.rptViewer.Name = "rptViewer"
        Me.rptViewer.ReportSource = Nothing
        Me.rptViewer.ShowGroupTreeButton = False
        Me.rptViewer.Size = New System.Drawing.Size(752, 480)
        Me.rptViewer.TabIndex = 0
        '
        'frmReportViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(768, 477)
        Me.Controls.Add(Me.rptViewer)
        Me.Name = "frmReportViewer"
        Me.Text = "frmReportViewer"
        Me.ResumeLayout(False)

    End Sub
    Private Sub frmReportViewer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        rptViewer.Height = Me.Height
        rptViewer.Width = Me.Width
        rptViewer.DisplayGroupTree = False
        rptViewer.ShowPrintButton = True
        rptViewer.RefreshReport()
    End Sub
    Private Sub frmReportViewer_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        rptViewer.Left = 8
        rptViewer.Top = 8
        rptViewer.Width = Me.Width - 16
        rptViewer.Height = Me.Height - 16
        rptViewer.Refresh()
        Me.Refresh()
    End Sub


    Public Sub GenerateReport(ByVal Title As String, ByVal SourceDataSet As System.Data.DataSet, ByVal RptClass As CrystalDecisions.CrystalReports.Engine.ReportClass)
        Me.Text = Title
        GenerateReport(SourceDataSet, RptClass)
    End Sub
    Public Sub GenerateReport(ByVal SourceDataSet As System.Data.DataSet, ByVal RptClass As CrystalDecisions.CrystalReports.Engine.ReportClass)
        Try
            RptClass.SetDataSource(SourceDataSet)
            rptViewer.ReportSource = RptClass

            rptViewer.Refresh()
        Catch ex As Exception
            Throw (ex)
        Finally
            RptClass = Nothing
        End Try
    End Sub


#End Region

    Private Sub rptViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rptViewer.Load

    End Sub
End Class
