Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BackgroundColor As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Font As System.Web.UI.WebControls.DropDownList
    Protected WithEvents SubmitButton As System.Web.UI.WebControls.Button
    Protected WithEvents txt As System.Web.UI.WebControls.TextBox

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Put user code to initialize the page here
        Dim sText As String = GetName(Request.Url.Query())
        Dim sFont As String = "ËÎÌו"
        Dim sID As String = Request.QueryString("id")
        Dim isExcellent As Boolean = Request.QueryString("isexcellent")
        Dim imageSrc As String = "images/cert_template.jpg"

        If isExcellent Then
            imageSrc = "images/cert_template2.jpg"
        End If

        Dim oBitmap As Drawing.Bitmap = New Drawing.Bitmap(Server.MapPath("images/cert_template.jpg"))
        Dim oGraphic As Drawing.Graphics = Drawing.Graphics.FromImage(oBitmap)
        Dim oColor As System.Drawing.Color

        Dim oBrushWrite As New Drawing.SolidBrush(Drawing.Color.Black)

        Dim oFont As New Drawing.Font(sFont, 20)
        Dim oPoint As New Drawing.PointF(385.0F, 278.0F)
        Dim oPoint1 As New Drawing.PointF(385.0F, 315.0F)

        'Response.Write("<script language='javascript'>alert('" + sText + "');</script>")

        oGraphic.DrawString(sText, oFont, oBrushWrite, oPoint)
        oGraphic.DrawString(sID, oFont, oBrushWrite, oPoint1)

        Response.ContentType = "image/jpeg"
        Response.AddHeader("Content-Disposition", "attachment; filename=certificate.jpg")

        oBitmap.Save(Response.OutputStream, Drawing.Imaging.ImageFormat.Jpeg)

    End Sub

    Public Function GetName(ByVal query As String) As String
        Dim result As String
        If Not query Is Nothing Then
            If query.StartsWith("?") Then
                query = query.Substring(1)
            End If

            Dim pair As String
            Dim sets = query.Split("&"c)
            For Each pair In sets
                Dim parts As Array = pair.Split("="c)
                If parts(0) = "name" Then
                    If parts.Length = 1 Then
                        Return String.Empty
                    Else
                        Return Web.HttpUtility.UrlDecode(parts(1), System.Text.Encoding.Default)
                    End If
                End If
            Next
        End If
        Return String.Empty
    End Function

End Class
