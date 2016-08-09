<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EPIExcel.aspx.cs" Inherits="Excel.EPIExcel" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="margin-top:50px">
    <asp:Label ID="Label1" runat="server" Text="ExcelExport:"></asp:Label>
    <asp:button runat="server" text="Excel" style="margin-left :30px" OnClick="ExcelButton_Click"/>
    </div>
    </form>
</body>
</html>
