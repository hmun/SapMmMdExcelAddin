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
        Me.GroupMaterialMaster = Me.Factory.CreateRibbonGroup
        Me.ButtonSapMaterialGetAll = Me.Factory.CreateRibbonButton
        Me.ButtonSapMaterialChange = Me.Factory.CreateRibbonButton
        Me.GroupMaterialPrice = Me.Factory.CreateRibbonGroup
        Me.ButtonSapMaterialPriceChange = Me.Factory.CreateRibbonButton
        Me.GroupSourceList = Me.Factory.CreateRibbonGroup
        Me.ButtonSapSourceListRead = Me.Factory.CreateRibbonButton
        Me.ButtonSapSourceListUpdate = Me.Factory.CreateRibbonButton
        Me.GroupRouting = Me.Factory.CreateRibbonGroup
        Me.ButtonSapRoutingCreate = Me.Factory.CreateRibbonButton
        Me.ButtonSapRoutingChange = Me.Factory.CreateRibbonButton
        Me.GroupSAPLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.GroupGoodsMovement = Me.Factory.CreateRibbonGroup
        Me.ButtonSapGoodsMovementCreate = Me.Factory.CreateRibbonButton
        Me.SapMmMd.SuspendLayout()
        Me.GroupMaterialMaster.SuspendLayout()
        Me.GroupMaterialPrice.SuspendLayout()
        Me.GroupSourceList.SuspendLayout()
        Me.GroupRouting.SuspendLayout()
        Me.GroupSAPLogon.SuspendLayout()
        Me.GroupGoodsMovement.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapMmMd
        '
        Me.SapMmMd.Groups.Add(Me.GroupMaterialMaster)
        Me.SapMmMd.Groups.Add(Me.GroupMaterialPrice)
        Me.SapMmMd.Groups.Add(Me.GroupSourceList)
        Me.SapMmMd.Groups.Add(Me.GroupRouting)
        Me.SapMmMd.Groups.Add(Me.GroupGoodsMovement)
        Me.SapMmMd.Groups.Add(Me.GroupSAPLogon)
        Me.SapMmMd.Label = "SAP MM Md"
        Me.SapMmMd.Name = "SapMmMd"
        '
        'GroupMaterialMaster
        '
        Me.GroupMaterialMaster.Items.Add(Me.ButtonSapMaterialGetAll)
        Me.GroupMaterialMaster.Items.Add(Me.ButtonSapMaterialChange)
        Me.GroupMaterialMaster.Label = "Material Master"
        Me.GroupMaterialMaster.Name = "GroupMaterialMaster"
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
        'GroupMaterialPrice
        '
        Me.GroupMaterialPrice.Items.Add(Me.ButtonSapMaterialPriceChange)
        Me.GroupMaterialPrice.Label = "Material Price"
        Me.GroupMaterialPrice.Name = "GroupMaterialPrice"
        '
        'ButtonSapMaterialPriceChange
        '
        Me.ButtonSapMaterialPriceChange.Image = CType(resources.GetObject("ButtonSapMaterialPriceChange.Image"), System.Drawing.Image)
        Me.ButtonSapMaterialPriceChange.Label = "Post Price Change"
        Me.ButtonSapMaterialPriceChange.Name = "ButtonSapMaterialPriceChange"
        Me.ButtonSapMaterialPriceChange.ShowImage = True
        '
        'GroupSourceList
        '
        Me.GroupSourceList.Items.Add(Me.ButtonSapSourceListRead)
        Me.GroupSourceList.Items.Add(Me.ButtonSapSourceListUpdate)
        Me.GroupSourceList.Label = "Source List"
        Me.GroupSourceList.Name = "GroupSourceList"
        '
        'ButtonSapSourceListRead
        '
        Me.ButtonSapSourceListRead.Image = CType(resources.GetObject("ButtonSapSourceListRead.Image"), System.Drawing.Image)
        Me.ButtonSapSourceListRead.Label = "Read Source List"
        Me.ButtonSapSourceListRead.Name = "ButtonSapSourceListRead"
        Me.ButtonSapSourceListRead.ShowImage = True
        '
        'ButtonSapSourceListUpdate
        '
        Me.ButtonSapSourceListUpdate.Image = CType(resources.GetObject("ButtonSapSourceListUpdate.Image"), System.Drawing.Image)
        Me.ButtonSapSourceListUpdate.Label = "Update Source List"
        Me.ButtonSapSourceListUpdate.Name = "ButtonSapSourceListUpdate"
        Me.ButtonSapSourceListUpdate.ShowImage = True
        '
        'GroupRouting
        '
        Me.GroupRouting.Items.Add(Me.ButtonSapRoutingCreate)
        Me.GroupRouting.Items.Add(Me.ButtonSapRoutingChange)
        Me.GroupRouting.Label = "Routing"
        Me.GroupRouting.Name = "GroupRouting"
        '
        'ButtonSapRoutingCreate
        '
        Me.ButtonSapRoutingCreate.Image = CType(resources.GetObject("ButtonSapRoutingCreate.Image"), System.Drawing.Image)
        Me.ButtonSapRoutingCreate.Label = "Create Routing"
        Me.ButtonSapRoutingCreate.Name = "ButtonSapRoutingCreate"
        Me.ButtonSapRoutingCreate.ShowImage = True
        '
        'ButtonSapRoutingChange
        '
        Me.ButtonSapRoutingChange.Image = CType(resources.GetObject("ButtonSapRoutingChange.Image"), System.Drawing.Image)
        Me.ButtonSapRoutingChange.Label = "Change Routing"
        Me.ButtonSapRoutingChange.Name = "ButtonSapRoutingChange"
        Me.ButtonSapRoutingChange.ShowImage = True
        '
        'GroupSAPLogon
        '
        Me.GroupSAPLogon.Items.Add(Me.ButtonLogon)
        Me.GroupSAPLogon.Items.Add(Me.ButtonLogoff)
        Me.GroupSAPLogon.Label = "SAP Logon"
        Me.GroupSAPLogon.Name = "GroupSAPLogon"
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
        'GroupGoodsMovement
        '
        Me.GroupGoodsMovement.Items.Add(Me.ButtonSapGoodsMovementCreate)
        Me.GroupGoodsMovement.Label = "Goods Movement"
        Me.GroupGoodsMovement.Name = "GroupGoodsMovement"
        '
        'ButtonSapGoodsMovementCreate
        '
        Me.ButtonSapGoodsMovementCreate.Image = CType(resources.GetObject("ButtonSapGoodsMovementCreate.Image"), System.Drawing.Image)
        Me.ButtonSapGoodsMovementCreate.Label = "Create Goods Movement"
        Me.ButtonSapGoodsMovementCreate.Name = "ButtonSapGoodsMovementCreate"
        Me.ButtonSapGoodsMovementCreate.ShowImage = True
        '
        'SapMmMdRibbon
        '
        Me.Name = "SapMmMdRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapMmMd)
        Me.SapMmMd.ResumeLayout(False)
        Me.SapMmMd.PerformLayout()
        Me.GroupMaterialMaster.ResumeLayout(False)
        Me.GroupMaterialMaster.PerformLayout()
        Me.GroupMaterialPrice.ResumeLayout(False)
        Me.GroupMaterialPrice.PerformLayout()
        Me.GroupSourceList.ResumeLayout(False)
        Me.GroupSourceList.PerformLayout()
        Me.GroupRouting.ResumeLayout(False)
        Me.GroupRouting.PerformLayout()
        Me.GroupSAPLogon.ResumeLayout(False)
        Me.GroupSAPLogon.PerformLayout()
        Me.GroupGoodsMovement.ResumeLayout(False)
        Me.GroupGoodsMovement.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapMmMd As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GroupMaterialMaster As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GroupSAPLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapMaterialChange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapMaterialGetAll As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupMaterialPrice As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapMaterialPriceChange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupSourceList As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapSourceListRead As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapSourceListUpdate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupRouting As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapRoutingCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapRoutingChange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupGoodsMovement As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapGoodsMovementCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapMmMdRibbon
        Get
            Return Me.GetRibbon(Of SapMmMdRibbon)()
        End Get
    End Property
End Class
