<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="project.WebForm1" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <!-- CSS only -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65" crossorigin="anonymous">
    <style type="text/css">
        .style{
           
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
        .myGridClass th {
            padding: 4px 2px;
            color: #fff;
            background: #424242;
            border-left: solid 1px #525252;
            font-size: 0.9em;
        }
        .myGridClass td {
            padding: 2px;
            border: solid 1px #c1c1c1;
            
        }
        .myGridClass {
            width: 100%;
            /*this will be the color of the odd row*/
            background-color: #fff;
            margin: 5px 0 10px 0;
            border: solid 1px #525252;
            border-collapse:collapse;

        }
        .right {
            text-align: right;
            
         }
        .blue2 {
            background-color: #3399FF
        }
        .white {
            color:white
        }
        .bg {
            position: fixed;
            top: 0;
            left: 0;
            bottom: 0;
            right: 0;
            z-index: -999;
        }
        .bg img {
            min-height: 100%;
            min-width: 1000px;
            width: 100%;
        }

        @media screen and (max-width: 1000px) {
            img.bg {
                left: 50%;
                margin-left: -500px;
            }
        }
     </style>
    <script>
        

        function alert(message,type) {
            var wrapper = document.createElement('div');
            var alertPlaceholder = document.getElementById('liveAlertPlaceholder');
            wrapper.innerHTML = '<div class="alert alert-' + type + ' alert-dismissible text-center text-success" role="alert">' + message + '<button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button></div>';

            alertPlaceholder.append(wrapper);
        }
    </script>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-F3w7mX95PdgyTmZZMECAngseQB83DfGTowi0iMjiWaeVhAn4FJkqJByhZMI3AhiU" crossorigin="anonymous">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.1/dist/js/bootstrap.bundle.min.js" integrity="sha384-/bQdsTh/da6pkI1MST/rWKFNjaCP5gBSY4sEBT38Q/9RBh9AH40zEOg7Hlq2THRZ" crossorigin="anonymous"></script>
</head>
<body>
    
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" />
        <div id="liveAlertPlaceholder"></div>
        <div
        class="bg"
        style="background-color:white;">
            <img src="https://png.pngtree.com/background/20220725/original/pngtree-delivery-cartoon-picture-image_1752746.jpg"/>
        </div>
        <h1 style="background-color:MediumSeaGreen;text-align:center; color:white;">行得利貨物對比系統</h1>
        <div style="text-align:center;color:red; font-weight:bold;">
            <br />
            <asp:Label ID="Label5" runat="server" Text="" CssClass="white" ></asp:Label>
        </div>
        <br />
        <div class="container-fluid">
            <div class="row">
                <div class="col-md-3 offset-md-1 bg-info text-dark border border-primary">
                    <div class="row-auto">
                    <p>選擇檔案</p> <asp:FileUpload ID="FileUpload1" runat="server"/>   
                    </div>
                    
                    <div class="row">              
                        <div class="co1 right">
                            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="匯入" CssClass="btn btn-primary mx-5"/>
                        </div>                    
                         <asp:Label ID="Label1" runat="server" Text="" CssClass="right"></asp:Label>                       
                    </div>
                                     
                </div>
                <div class="col-md-3 bg-info text-dark border-top border-bottom border-end border-primary">
                    <div class="row-auto">
                        <p>選擇對比檔案</p>
                        <asp:FileUpload ID="FileUpload2" runat="server" CssClass=""/>  
                    </div>
                    
                    <div class="row">
                        <div class="co1 right">
                            <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="匯入" CssClass="btn btn-primary me-3"/>  
                        </div>
                        <asp:Label ID="Label2" runat="server" Text="" CssClass="right"></asp:Label>
                    </div>
                </div>

                <div class="col-md-3 offset-sm-1">
                    <div class="row-auto gy-5">
                        <asp:RadioButton ID="RadioButton1" runat="server"  Text="比對相同" GroupName="a"/>
                        <asp:RadioButton ID="RadioButton2" runat="server"  Text="比對不相同" GroupName="a"/>                                   
                        <asp:RadioButton ID="RadioButton3" runat="server"  Text="合併" GroupName="a"/>
                    </div>
                    <div class="row-auto">
                        <asp:Button ID="Button3" runat="server" OnClick="Button3_Click" Text="執行並下載" CssClass="style mt-2 ms-5"/> 
                    </div>
                </div>
           
            </div>
        <br/> 
        </div>
        <div class="container-fluid">
            <div class="row-auto">
                <asp:Button ID="Button4" runat="server" OnClick="Button4_Click" CssClass="Button Button1" Text="清除所有資料"/>
                <asp:Button ID="Button5" runat="server" OnClick="Button5_Click" CssClass="Button Button2" Text="清除對比資料"/>
                <br/>
                <br/>
                <br/>

                
            </div>
        </div>
        <br/> 
        <br/>
        <br/>
        <div class="container-fluid">
            <div class="col-md-3 offset-md-1">
                <div class="card" style="width: 30rem;">
                    <div class="card-body">
                        <ul>
                            <li>比對相同、不相同: 先放多欄位，再放少欄位</li>
                            <li>合併: 主資料:a,b,c 對比資料: a,b,c,d 合併</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
        <br/>
        <br/>
        <br/>
        <br/>
        <br/>
        <br/>
             
        <div class="row">
            <div class="col-md-3 offset-md-1">
                <asp:GridView ID="GridView1" runat="server" CssClass="myGridClass">
                </asp:GridView>
            </div>
            <div class="col-md-3 offset-md-2">
                <asp:GridView ID="GridView2" runat="server" CssClass="myGridClass">
                </asp:GridView>  
            </div>
        </div>
          
            
        
        <br/>
        <br/>     
        <br/>
        <br/>
        
    
    </form>
</body>
</html>
