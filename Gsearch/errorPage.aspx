<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="errorPage.aspx.cs" Inherits="errorPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div>
            <h3 class="font_black font_size15">錯誤訊息</h3>
            <table>
                <tr>
                    <td class="font_black font_size15">
                        <br />
                        <%
                            string ErrorMsg = (Application["ErrorMsg"] == null) ? "您不可以進入本網頁!" : (string)Application["ErrorMsg"];
                            Response.Write(ErrorMsg);
                        %>                        
                    </td>
                </tr>
                <tr>
                    <td class="font_size13">
                    <a href="<%=ResolveUrl("~/index.aspx")%>" target="_self">返回首頁</a>
                    </td>
                </tr>
             </table>
    </div>
</asp:Content>

