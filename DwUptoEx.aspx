<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="DwUptoEx.aspx.cs" Inherits="DWStock.DwUptoEx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <asp:Button Text="Upload" OnClick="Upload" runat="server" />
    </form>
    <script type="text/javascript">  
        $(document).ready(function () {
            $("#GridView1").prepend($("<thead></thead>").append($(this).find("tr:first"))).dataTable();
        });
    </script>
</body>
</html>
