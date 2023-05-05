Partial Class SapMmMdRibbon
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapMmMdRibbon))
        Me.SapMmMd = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ButtonSapMaterialGetAll = Me.Factory.CreateRibbonButton
        Me.ButtonSapMaterialChange = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.ButtonSapMaterialPriceChange = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapMmMd.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapMmMd
        '
        Me.SapMmMd.Groups.Add(Me.Group1)
        Me.SapMmMd.Groups.Add(Me.Group3)
        Me.SapMmMd.Groups.Add(Me.Group2)
        Me.SapMmMd.Label = "SAP MM Md"
        Me.SapMmMd.Name = "SapMmMd"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ButtonSapMaterialGetAll)
        Me.Group1.Items.Add(Me.ButtonSapMaterialChange)
        Me.Group1.Label = "Material Master"
        Me.Group1.Name = "Group1"
        '
        'ButtonSapMaterialGetAll
        '
        Me.ButtonSapMaterialGetAll.Image = CType(resources.GetObject("ButtonSapMaterialGetAll.Image"), System.Drawing.Image)
        Me.ButtonSapMaterialGetAll.Label = "Read Material"
        Me.ButtonSapMaterialGetAll.Name = "ButtonSapMaterialGetAll"
        Me.ButtonSapMaterialGetAll.ShowImage = True
        '
        'ButtonSapMaterialChange
        '
        Me.ButtonSapMaterialChange.Image = CType(resources.GetObject("ButtonSapMaterialChange.Image"), System.Drawing.Image)
        Me.ButtonSapMaterialChange.Label = "Change Material"
        Me.ButtonSapMaterialChange.Name = "ButtonSapMaterialChange"
        Me.ButtonSapMaterialChange.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ButtonSapMaterialPriceChange)
        Me.Group3.Label = "Material Price"
        Me.Group3.Name = "Group3"
        '
        'ButtonSapMaterialPriceChange
        '
        Me.ButtonSapMaterialPriceChange.Image = CType(resources.GetObject("ButtonSapMaterialPriceChange.Image"), System.Drawing.Image)
        Me.ButtonSapMaterialPriceChange.Label = "Post Price Change"
        Me.ButtonSapMaterialPriceChange.Name = "ButtonSapMaterialPriceChange"
        Me.ButtonSapMaterialPriceChange.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ButtonLogon)
        Me.Group2.Items.Add(Me.ButtonLogoff)
        Me.Group2.Label = "SAP Logon"
        Me.Group2.Name = "Group2"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'SapMmMdRibbon
        '
        Me.Name = "SapMmMdRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapMmMd)
        Me.SapMmMd.ResumeLayout(False)
        Me.SapMmMd.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapMmMd As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapMaterialChange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapMaterialGetAll As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapMaterialPriceChange As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapMmMdRibbon
        Get
            Return Me.GetRibbon(Of SapMmMdRibbon)()
        End Get
    End Property
End Class
