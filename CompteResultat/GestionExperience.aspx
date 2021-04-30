<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="GestionExperience.aspx.cs" Inherits="CompteResultat.GestionExperience" EnableViewState="false"%>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
    
    <style>
        input[type="submit"]:disabled {background: #dddddd;}

    </style>
    <script type="text/javascript">

        $(document).ready(function () { 
            $("#cmdImport").click(function (evt) {
                $("#divLoading").css("display", "block");
            });  
        }); 
     
   </script>

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

    <div style="float: left; width: 400px;">
        <h1><asp:Literal  ID="Literal9" runat="server">Liste des assureurs :</asp:Literal> </h1> 
        
        <asp:CheckBox ID="chkAssur" runat="server" Checked="false" AutoPostBack="true"
              style="display:inline-block;margin-bottom:15px; margin-top:15px;" Text="Afficher uniquement les assureurs sans experience" /><br />  

        <asp:ListBox ID="lbAssur" runat="server" SelectMethod="GetAssureurs" DataTextField="Name" DataValueField="Id" Height="250px" Width="100%" 
                AutoPostBack="true" OnSelectedIndexChanged="lbAssur_SelectedIndexChanged" EnableViewState="false" OnDataBound="lbAssur_DataBound" >            
        </asp:ListBox>

        <div style="display:none;"> 
            <asp:FileUpload Width="150" ID="uploadExcel" runat="server" onchange="this.form.submit()" /> 
        </div>
        <div style="margin-top:15px; float:right;">
            <asp:Button CssClass="ButtonBigBlue ButtonInline" style="margin-right:20px; width:90px;" ID="cmdImport" runat="server" Text="Import" ClientIDMode="Static" />  
            <asp:Button CssClass="ButtonBigBlue ButtonInline" style="margin-right:20px; width:90px;" ID="cmdExport" runat="server" Text="Export" OnClick="cmdExport_Click"  /> 
            <asp:Button CssClass="ButtonBigRed ButtonInline" style="width:90px;" ID="cmdDelete" runat="server" Text="Supprimer" OnClick="cmdDelete_Click" ClientIDMode="Static"  /> 
                         
            <%-- <input type="submit" ID="cmdDelete" name="cmdDelete" class="ButtonBigRed ButtonInline" style="width:90px;" value="Supprimer" runat="server" />  --%>
        </div>

        <div style="margin-top:5px; float:right;">
            <asp:Button CssClass="ButtonBigBlue ButtonInline" style="margin-right:0px; width:320px;" ID="cmdRecreate" runat="server" Text="REGENERER à partir des PRESTATIONS" ClientIDMode="Static" OnClick="cmdRecreate_Click" />  
        </div>

        <div runat="server" id="divLoading"  style="display:none" ClientIDMode="Static" >
            <img width="100px" height="100px" style="margin: 20px 5px 10px 140px;" src="Images/ajax-loader.gif" />
        </div>

        <asp:ValidationSummary style="margin-top:15px; float:left;" ForeColor="Red" ID="ValSummary" runat="server" />  

    </div>


    <%-- Repeater --%>

    <div style="float: left; margin-left:80px; "> 

        <h1><asp:Literal  ID="Literal1" runat="server">Experience :</asp:Literal> </h1> 
        
        <div class="Repeater"> 

        <asp:PlaceHolder ID="phHeader" Visible='false' runat="server">
            <asp:Label ID="lblEmpty" runat="server" Text="Il n'y a pas des données disponibles !"> </asp:Label>   
        </asp:PlaceHolder>

        <asp:Repeater ItemType="CompteResultat.DAL.C_TempExpData" SelectMethod="GetExperience" ID="rptExp" runat="server" OnItemDataBound="rptExp_ItemDataBound">
            <HeaderTemplate>      
                  <table>
                      <tr><th>Année</th><th>Contrat</th><th>Acte</th><th>Nombre Acte</th><th>Frais Réel</th></tr>                    
            </HeaderTemplate>
            <FooterTemplate>
                </table>                
            </FooterTemplate>

            <ItemTemplate>
                <tr>
                    <td><%#: Item.AnneeExp %></td>
                    <td><%#: Item.Contrat %></td>
                    <td><%#: Item.LibActe %></td>
                    <td><%#: Item.NombreActe %></td>
                    <td><%#: Item.Fraisreel.Value.ToString("F2") %></td>
                </tr>
            </ItemTemplate>
            <AlternatingItemTemplate>
                <tr class="alternate">
                    <td><%#: Item.AnneeExp %></td>
                    <td><%#: Item.Contrat %></td>
                    <td><%#: Item.LibActe %></td>
                    <td><%#: Item.NombreActe %></td>
                    <td><%#: Item.Fraisreel.Value.ToString("F2") %></td>
                </tr>
            </AlternatingItemTemplate>
            
        </asp:Repeater>
            
        </div>     

    </div>

</asp:Content>

