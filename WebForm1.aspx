<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="project.WebForm1" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .style{
           
            position:absolute;
            right:30%;
            color:darkorchid;
        }
        .Button
       {
        display:flex;
       	background-color: aqua;
        border: none;
        color: white;
        padding: 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        color:black;
        font-size: 16px;
        margin: 4px 2px;       
        border-radius: 50%; 
        
        }
        .Button1{position:absolute; right:10%;}
        .Button2{position:absolute; right:20%;}
        </style>
</head>
<body>
    <form id="form1" runat="server">
        <h1 style="background-color:MediumSeaGreen;text-align:center; color:white;">行得利貨物對比系統</h1>
        <div style="text-align:center;color:red; font-weight:bold;">
            <asp:Label ID="Label5" runat="server" Text="" ></asp:Label>
        </div>
        <br />
        <div>
            選擇檔案&nbsp;  <asp:FileUpload ID="FileUpload1" runat="server" Height="26px" Width="263px" />   
            <asp:Label ID="Label3" runat="server" Text="選擇對比檔案" style="position:absolute; left:35%;"></asp:Label><asp:FileUpload ID="FileUpload2" runat="server" Height="26px" Width="263px" style="position:absolute; left:42%;" />  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;          
            <asp:RadioButton ID="RadioButton1" runat="server" style="position:absolute; right:30%;" Text="比對相同" GroupName="a"/>
            <asp:RadioButton ID="RadioButton2" runat="server" style="position:absolute; right:22%;" Text="比對不相同" GroupName="a"/>
            <asp:RadioButton ID="RadioButton3" runat="server" style="position:absolute; right:17%;" Text="合併" GroupName="a"/>
        <br/> 
        </div>
        &nbsp
        <asp:panel ID="panel1" runat="server" >
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="匯入"/>&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label1" runat="server" Text="" style="position:absolute; left:5%;"></asp:Label>           
            <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="匯入" style="position:absolute; left:35%;"/> &nbsp;&nbsp;&nbsp;                       
            <asp:Label ID="Label2" runat="server" Text="" style="position:absolute; left:40%;"></asp:Label>
            <asp:Button ID="Button3" runat="server" OnClick="Button3_Click" Text="執行並下載" CssClass="style"/> 
            <br/>
            <br/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        </asp:panel>
            <asp:Button ID="Button4" runat="server" OnClick="Button4_Click" CssClass="Button Button1" Text="清除所有資料"/>
            <asp:Button ID="Button5" runat="server" OnClick="Button5_Click" CssClass="Button Button2" Text="清除對比資料"/>
        <asp:Label ID="Label4" runat="server" Text="" style="position:absolute; right:15%; top:22%; color:red"></asp:Label>
        <br/>
        <br />
        <br />
        <br/>
        <asp:panel ID="panel2" runat="server">
            <asp:GridView ID="GridView1" runat="server" >
            </asp:GridView>
            <asp:GridView ID="GridView2" runat="server" >
            </asp:GridView>    
            
        </asp:panel>
        <br/>
        <br/>     
        <br/>
        <br/>
      
    </form>
</body>
</html>
