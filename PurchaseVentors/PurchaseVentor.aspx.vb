Public Class PurchaseVentor
    Inherits System.Web.UI.Page
    Dim Sqlconn As String = ConfigurationManager.ConnectionStrings("conString35").ConnectionString
    '  Dim Sqlconn As String = ConfigurationManager.ConnectionStrings("conStringFamily").ConnectionString
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BuildVendorChange()
        End If
    End Sub


    Sub BuildVendorChange()
        Dim index As Int32
        Dim ds As DataSet
        Dim dt = New DataBaseProc
        index = ddVendorNameUpdate.SelectedIndex
        ds = dt.GetPurchaseVentor(Sqlconn)
        ddVendorNameUpdate.DataSource = ds
        ddVendorNameUpdate.DataTextField = "VendorName"
        ddVendorNameUpdate.DataValueField = "RecordID"
        ddVendorNameUpdate.DataBind()
        ddVendorNameUpdate.SelectedIndex = index
    End Sub

    Protected Sub addTitle_Click1(sender As Object, e As EventArgs)
        Dim obj1 As New DataBaseProc
        Dim sql As New StringBuilder
        Dim obj As New Object

        obj = obj1.InsertPurchaseVentor(txtProgramTitle.Text, chk1.Checked, Sqlconn)
        BuildVendorChange()
    End Sub

    Protected Sub changeTitle_Click(sender As Object, e As EventArgs)
        Dim obj1 As New DataBaseProc
        Dim sql As New StringBuilder
        Dim obj As New Object
        Dim inactive As New Int32
        If chk1.Checked Then
            inactive = 1
        Else
            inactive = 0
        End If

        obj = obj1.UpdatePurchaseVentor(ddVendorNameUpdate.SelectedValue, txtVendorNameChg.Text, inactive.ToString(), Sqlconn)
        BuildVendorChange()
    End Sub

End Class