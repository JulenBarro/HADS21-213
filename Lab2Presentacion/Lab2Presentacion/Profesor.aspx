<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="Lab2Presentacion.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .auto-style1 {
            height: 23px;
        }
        .auto-style2 {
            height: 23px;
            width: 378px;
        }
        .auto-style3 {
            width: 100%;
            height: 592px;
        }
    </style>
</head>
<body style="height: 637px">
    <form id="form1" runat="server">
    <table class="auto-style3">
        <tr>
            <td class="auto-style2" style="background-color:papayawhip; border:outset" >
                <div>
                <label>Asignaturas</label>
                </div>
                <hr/>
                <div>
                <label><a href='TareasProfesor.aspx'>Tareas</a></label>
                </div>
                <hr/>
                <div>
                <label>Grupos</label>
                </div>
                <hr/>
                <div>
                    <asp:Panel ID="Panel1" runat="server">
                        <asp:LinkButton ID="LinkButton1" runat="server" PostBackUrl="~/ImportarXMLDocument.aspx">Importar v. XMLDocument</asp:LinkButton>
                    </asp:Panel>
                </div>
                <hr/>
                <div>
                    <asp:LinkButton ID="LinkButton2" runat="server" PostBackUrl="~/ExportarTarea.aspx">Exportar</asp:LinkButton>
                </div>
                <hr/>
                <div>
                <label>Importar v. DataSet</label>
                </div>
            </td>
            <td class="auto-style1" style="text-align: center; background-color: azure; border:outset">
                <label>Gestion web de tareas de dedicacion</label>
                <hr/>
                <label>Profesores</label>

            </td>
            
        </tr>
    </table>
    </form>
</body>
</html>
