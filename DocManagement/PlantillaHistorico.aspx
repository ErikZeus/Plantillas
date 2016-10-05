<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PlantillaHistorico.aspx.cs" Inherits="DocManagement.PlantillaHistorico" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Unity Promotores Plantilla</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="C#">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <style>
       .datagrid table { border-collapse: collapse; text-align: left; width: 100%; } 
       .datagrid {font: normal 16px/150% Arial, Helvetica, sans-serif; background: #fff; overflow: hidden; border: 1px solid #652299; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; }
       .datagrid table td, .datagrid table th { padding: 3px 10px; }
        .datagrid table td a { font-size:large;} 
       .datagrid table thead th {background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #652299), color-stop(1, #4D1A75) );background:-moz-linear-gradient( center top, #652299 5%, #4D1A75 100% );filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#652299', endColorstr='#4D1A75');background-color:#652299; color:#FFFFFF; font-size: 15px; font-weight: bold; border-left: 1px solid #714399; } 
       .datagrid table thead th:first-child { border: none; }
       .datagrid table tbody td { color: #4D1A75; border-left: 1px solid #E7BDFF;font-size: 12px;font-weight: normal; }
       .datagrid table tbody .alt td { background: #F4E3FF; color: #4D1A75; }
       .datagrid table tbody td:first-child { border-left: none; }
       .datagrid table tbody tr:last-child td { border-bottom: none; }
       .datagrid table tfoot td div { border-top: 1px solid #652299;background: #F4E3FF;} 
       .datagrid table tfoot td { padding: 0; font-size: 16px } 
       .datagrid table tfoot td div{ padding: 2px; }
       .datagrid table tfoot td ul { margin: 0; padding:0; list-style: none; text-align: right; }
       .datagrid table tfoot  li { display: inline; }
       .datagrid table tfoot li a { text-decoration: none; display: inline-block;  padding: 2px 8px; margin: 1px;color: #FFFFFF;border: 1px solid #652299;-webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #652299), color-stop(1, #4D1A75) );background:-moz-linear-gradient( center top, #652299 5%, #4D1A75 100% );filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#652299', endColorstr='#4D1A75');background-color:#652299; }
       .datagrid table tfoot ul.active, 
       .datagrid table tfoot ul a:hover { text-decoration: none;border-color: #4D1A75; color: #FFFFFF; background: none; background-color:#652299;}
       div.dhtmlx_window_active, div.dhx_modal_cover_dv { position: fixed !important; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div style="float: left;">

            <h1>Unity Promotores Plantillas 
            </h1>
            <h1>Historico
            </h1>
            <h1><asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            </h1>

            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                <ContentTemplate>
                    <div>
                        <a>Bajando Archivo Id Cliente : </a>
                        <asp:Label ID="lblCliente" runat="server" Text="."></asp:Label>
                    </div>
                    <div><a>Haga click en la plantilla que va a utilizar.</a></div>
                    <div class="datagrid">
                        <asp:Table id="Table1" BorderWidth="1" GridLines="Both" runat="server" Width="100%" Font-Bold="True" />
                    </div>
 

                </ContentTemplate>
            </asp:UpdatePanel>

        </div>
    </form>
</body>
</html>
