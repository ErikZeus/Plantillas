<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Llamada.aspx.cs" Inherits="DocManagement.Llamada" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
        <script type="text/javascript" language="javascript">
    function colorChanged(sender) 
    {
        sender.get_element().style.backcolor = 
        "#" + sender.get_selectedColor();
    }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
         <table bgcolor="White">
            <tr>
                <td style="background-color: #FFFFCC; border: 1px solid #000000">
                    <asp:UpdatePanel ID="panelInputs" runat="server">
                        <ContentTemplate>
                            <table>
                                <tr>
                                    <td colspan="2">
                                        <asp:TextBox ID="txtData" runat="server" Width="264px">038000356216</asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style1" align="center">
                                        <span lang="en-us">Encoding</span></td>
                                    <td align="center">
                                        <span lang="en-us">Image Format</span></td>
                                </tr>
                                <tr>
                                    <td class="style1">
                                        <asp:Button ID="btnEncode" runat="server" Text="Encode" Width="100px" 
                                            onclick="btnEncode_Click" UseSubmitBehavior="False" />
                                    </td>
                                    <td>
                                        &nbsp;</td>
                                </tr>
                            </table>
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </td>
                <td style="border-color: #000000; border-width: 1px; border-top-style: solid; border-bottom-style: solid">
                    &nbsp;</td>
                <td style="border-color: #000000; border-width: 1px; border-top-style: solid; border-bottom-style: solid">
                    <asp:UpdatePanel ID="panelBarcode" runat="server">
                        <ContentTemplate>
                            <asp:Image ID="BarcodeImage" runat="server" />
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </td>
                <td style="border-color: #000000; border-width: 1px; border-top-style: solid; border-bottom-style: solid; border-right-style: solid;">
                    &nbsp;</td>
            </tr>
            </table>
        <br />
    
    </div>
    </form>
</body>
</html>
