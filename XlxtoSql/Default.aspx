<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="XlxtoSql.Default" %>

<!DOCTYPE html  PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>DEMO FOR UPLOAD A XLX FILE INTO SQL SERVER</title>
</head>
<body bgcolor="darkturquoise">
    <form id="form1" runat="server">
        <div style="color:white">
            <table>  
            <tr>  
                <td>  
                    Select File  
                </td>  
                <td>  
                    <asp:FileUpload ID="FileUpload1" runat="server" />  
                </td>  
                <td>  
                </td>  
                <td>  
                    <asp:Button ID="Button1" runat="server" Text="Upload" OnClick="Button1_Click" />  
                </td>  
            </tr>  
        </table>  
        </div>
    </form>
</body>
</html>
