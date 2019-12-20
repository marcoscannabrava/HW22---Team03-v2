Partial Class SpartanTraderRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SpartanTraderRbn = Me.Factory.CreateRibbonTab
        Me.DatabaseGroup = Me.Factory.CreateRibbonGroup
        Me.AlphaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.BetaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.DashboardGroup = Me.Factory.CreateRibbonGroup
        Me.DashboardBtn = Me.Factory.CreateRibbonButton
        Me.MarketsGroup = Me.Factory.CreateRibbonGroup
        Me.StocksBtn = Me.Factory.CreateRibbonButton
        Me.OptionsBtn = Me.Factory.CreateRibbonButton
        Me.IndexBtn = Me.Factory.CreateRibbonButton
        Me.AutomationGroup = Me.Factory.CreateRibbonGroup
        Me.MainTimerToggleBtn = Me.Factory.CreateRibbonToggleButton
        Me.AutoPilotTglBtn = Me.Factory.CreateRibbonToggleButton
        Me.PortfolioGroup = Me.Factory.CreateRibbonGroup
        Me.InitialPositionBtn = Me.Factory.CreateRibbonButton
        Me.PortfolioBtn = Me.Factory.CreateRibbonButton
        Me.ConfirmTicketBtn = Me.Factory.CreateRibbonButton
        Me.PFResetBtn = Me.Factory.CreateRibbonButton
        Me.PFUpdateBtn = Me.Factory.CreateRibbonButton
        Me.EconomyGroup = Me.Factory.CreateRibbonGroup
        Me.EnvironmentBtn = Me.Factory.CreateRibbonButton
        Me.TCostsBtn = Me.Factory.CreateRibbonButton
        Me.ChartsGroup = Me.Factory.CreateRibbonGroup
        Me.StockOptionsBtn = Me.Factory.CreateRibbonButton
        Me.ProjectPFBtn = Me.Factory.CreateRibbonButton
        Me.ControlGroup = Me.Factory.CreateRibbonGroup
        Me.QuitBtn = Me.Factory.CreateRibbonButton
        Me.OverrideBox = Me.Factory.CreateRibbonEditBox
        Me.DateOverrideBtn = Me.Factory.CreateRibbonButton
        Me.CalcHedgeBtn = Me.Factory.CreateRibbonButton
        Me.SpartanTraderRbn.SuspendLayout()
        Me.DatabaseGroup.SuspendLayout()
        Me.DashboardGroup.SuspendLayout()
        Me.MarketsGroup.SuspendLayout()
        Me.AutomationGroup.SuspendLayout()
        Me.PortfolioGroup.SuspendLayout()
        Me.EconomyGroup.SuspendLayout()
        Me.ChartsGroup.SuspendLayout()
        Me.ControlGroup.SuspendLayout()
        '
        'SpartanTraderRbn
        '
        Me.SpartanTraderRbn.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.SpartanTraderRbn.Groups.Add(Me.DatabaseGroup)
        Me.SpartanTraderRbn.Groups.Add(Me.DashboardGroup)
        Me.SpartanTraderRbn.Groups.Add(Me.MarketsGroup)
        Me.SpartanTraderRbn.Groups.Add(Me.AutomationGroup)
        Me.SpartanTraderRbn.Groups.Add(Me.PortfolioGroup)
        Me.SpartanTraderRbn.Groups.Add(Me.EconomyGroup)
        Me.SpartanTraderRbn.Groups.Add(Me.ChartsGroup)
        Me.SpartanTraderRbn.Groups.Add(Me.ControlGroup)
        Me.SpartanTraderRbn.Label = "Spartan Trader"
        Me.SpartanTraderRbn.Name = "SpartanTraderRbn"
        '
        'DatabaseGroup
        '
        Me.DatabaseGroup.Items.Add(Me.AlphaTBtn)
        Me.DatabaseGroup.Items.Add(Me.BetaTBtn)
        Me.DatabaseGroup.Label = "DATABASE"
        Me.DatabaseGroup.Name = "DatabaseGroup"
        '
        'AlphaTBtn
        '
        Me.AlphaTBtn.Label = "Alpha"
        Me.AlphaTBtn.Name = "AlphaTBtn"
        '
        'BetaTBtn
        '
        Me.BetaTBtn.Label = "Beta"
        Me.BetaTBtn.Name = "BetaTBtn"
        '
        'DashboardGroup
        '
        Me.DashboardGroup.Items.Add(Me.DashboardBtn)
        Me.DashboardGroup.Label = "DASHBOARD"
        Me.DashboardGroup.Name = "DashboardGroup"
        '
        'DashboardBtn
        '
        Me.DashboardBtn.Label = "Dashboard"
        Me.DashboardBtn.Name = "DashboardBtn"
        '
        'MarketsGroup
        '
        Me.MarketsGroup.Items.Add(Me.StocksBtn)
        Me.MarketsGroup.Items.Add(Me.OptionsBtn)
        Me.MarketsGroup.Items.Add(Me.IndexBtn)
        Me.MarketsGroup.Label = "MARKETS"
        Me.MarketsGroup.Name = "MarketsGroup"
        '
        'StocksBtn
        '
        Me.StocksBtn.Label = "Stocks"
        Me.StocksBtn.Name = "StocksBtn"
        '
        'OptionsBtn
        '
        Me.OptionsBtn.Label = "Options"
        Me.OptionsBtn.Name = "OptionsBtn"
        '
        'IndexBtn
        '
        Me.IndexBtn.Label = "Index"
        Me.IndexBtn.Name = "IndexBtn"
        '
        'AutomationGroup
        '
        Me.AutomationGroup.Items.Add(Me.MainTimerToggleBtn)
        Me.AutomationGroup.Items.Add(Me.AutoPilotTglBtn)
        Me.AutomationGroup.Items.Add(Me.CalcHedgeBtn)
        Me.AutomationGroup.Label = "AUTOMATION"
        Me.AutomationGroup.Name = "AutomationGroup"
        '
        'MainTimerToggleBtn
        '
        Me.MainTimerToggleBtn.Label = "Engaged"
        Me.MainTimerToggleBtn.Name = "MainTimerToggleBtn"
        '
        'AutoPilotTglBtn
        '
        Me.AutoPilotTglBtn.Label = "Autopilot"
        Me.AutoPilotTglBtn.Name = "AutoPilotTglBtn"
        '
        'PortfolioGroup
        '
        Me.PortfolioGroup.Items.Add(Me.InitialPositionBtn)
        Me.PortfolioGroup.Items.Add(Me.PortfolioBtn)
        Me.PortfolioGroup.Items.Add(Me.ConfirmTicketBtn)
        Me.PortfolioGroup.Items.Add(Me.PFResetBtn)
        Me.PortfolioGroup.Items.Add(Me.PFUpdateBtn)
        Me.PortfolioGroup.Label = "PORTFOLIO"
        Me.PortfolioGroup.Name = "PortfolioGroup"
        '
        'InitialPositionBtn
        '
        Me.InitialPositionBtn.Label = "Initial Position"
        Me.InitialPositionBtn.Name = "InitialPositionBtn"
        '
        'PortfolioBtn
        '
        Me.PortfolioBtn.Label = "Portfolio"
        Me.PortfolioBtn.Name = "PortfolioBtn"
        '
        'ConfirmTicketBtn
        '
        Me.ConfirmTicketBtn.Label = "Confirmation Ticket"
        Me.ConfirmTicketBtn.Name = "ConfirmTicketBtn"
        '
        'PFResetBtn
        '
        Me.PFResetBtn.Label = "PF Reset"
        Me.PFResetBtn.Name = "PFResetBtn"
        '
        'PFUpdateBtn
        '
        Me.PFUpdateBtn.Label = "PF Update"
        Me.PFUpdateBtn.Name = "PFUpdateBtn"
        '
        'EconomyGroup
        '
        Me.EconomyGroup.Items.Add(Me.EnvironmentBtn)
        Me.EconomyGroup.Items.Add(Me.TCostsBtn)
        Me.EconomyGroup.Label = "ECONOMY"
        Me.EconomyGroup.Name = "EconomyGroup"
        '
        'EnvironmentBtn
        '
        Me.EnvironmentBtn.Label = "Environment"
        Me.EnvironmentBtn.Name = "EnvironmentBtn"
        '
        'TCostsBtn
        '
        Me.TCostsBtn.Label = "T. Costs"
        Me.TCostsBtn.Name = "TCostsBtn"
        '
        'ChartsGroup
        '
        Me.ChartsGroup.Items.Add(Me.StockOptionsBtn)
        Me.ChartsGroup.Items.Add(Me.ProjectPFBtn)
        Me.ChartsGroup.Label = "CHARTS"
        Me.ChartsGroup.Name = "ChartsGroup"
        '
        'StockOptionsBtn
        '
        Me.StockOptionsBtn.Label = "Stock / Options"
        Me.StockOptionsBtn.Name = "StockOptionsBtn"
        '
        'ProjectPFBtn
        '
        Me.ProjectPFBtn.Label = "Project PF"
        Me.ProjectPFBtn.Name = "ProjectPFBtn"
        '
        'ControlGroup
        '
        Me.ControlGroup.Items.Add(Me.QuitBtn)
        Me.ControlGroup.Items.Add(Me.OverrideBox)
        Me.ControlGroup.Items.Add(Me.DateOverrideBtn)
        Me.ControlGroup.Label = "CONTROL"
        Me.ControlGroup.Name = "ControlGroup"
        '
        'QuitBtn
        '
        Me.QuitBtn.Label = "Quit"
        Me.QuitBtn.Name = "QuitBtn"
        '
        'OverrideBox
        '
        Me.OverrideBox.Label = "Enter Date"
        Me.OverrideBox.Name = "OverrideBox"
        Me.OverrideBox.Text = "mm/dd/yy"
        '
        'DateOverrideBtn
        '
        Me.DateOverrideBtn.Label = "Date Override"
        Me.DateOverrideBtn.Name = "DateOverrideBtn"
        '
        'CalcHedgeBtn
        '
        Me.CalcHedgeBtn.Label = "Calc Hedge"
        Me.CalcHedgeBtn.Name = "CalcHedgeBtn"
        '
        'SpartanTraderRibbon
        '
        Me.Name = "SpartanTraderRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SpartanTraderRbn)
        Me.SpartanTraderRbn.ResumeLayout(False)
        Me.SpartanTraderRbn.PerformLayout()
        Me.DatabaseGroup.ResumeLayout(False)
        Me.DatabaseGroup.PerformLayout()
        Me.DashboardGroup.ResumeLayout(False)
        Me.DashboardGroup.PerformLayout()
        Me.MarketsGroup.ResumeLayout(False)
        Me.MarketsGroup.PerformLayout()
        Me.AutomationGroup.ResumeLayout(False)
        Me.AutomationGroup.PerformLayout()
        Me.PortfolioGroup.ResumeLayout(False)
        Me.PortfolioGroup.PerformLayout()
        Me.EconomyGroup.ResumeLayout(False)
        Me.EconomyGroup.PerformLayout()
        Me.ChartsGroup.ResumeLayout(False)
        Me.ChartsGroup.PerformLayout()
        Me.ControlGroup.ResumeLayout(False)
        Me.ControlGroup.PerformLayout()

    End Sub

    Friend WithEvents SpartanTraderRbn As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents DatabaseGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents MarketsGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ControlGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AlphaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents BetaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents StocksBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OptionsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents IndexBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents QuitBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DashboardGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DashboardBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PortfolioGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents InitialPositionBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PortfolioBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ConfirmTicketBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EconomyGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EnvironmentBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TCostsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PFResetBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PFUpdateBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OverrideBox As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents DateOverrideBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ChartsGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents StockOptionsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ProjectPFBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AutomationGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents MainTimerToggleBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents AutoPilotTglBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents CalcHedgeBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SpartanTraderRibbon() As SpartanTraderRibbon
        Get
            Return Me.GetRibbon(Of SpartanTraderRibbon)()
        End Get
    End Property
End Class
