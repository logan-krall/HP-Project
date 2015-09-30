<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Upload.aspx.cs" Inherits="HP_Analytics_Project.Images.WebForm1" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1>Data Analytics Application</h1>
            </hgroup>
            <span style="color:#fff">This tool was written to explore different statistical model options in the pursuit of generating actionable data from large data sets. </span>
        </div>
    </section>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server" HorizontalAlign="Center">
    <link rel="stylesheet" href="http://localhost:58937/code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
    <script src="//code.jquery.com/jquery-1.10.2.js"></script>
    <script src="//code.jquery.com/ui/1.8.2/jquery-ui.js"></script>

    <div style="text-align:center">
        <h3>Missing Row Values</h3> 
        <asp:Table ID="Table2" runat="server" Visible="true" HorizontalAlign="Center" />
            <p></p><hr /><p></p>
        <h3>Data Characteristics</h3>
        <asp:Table ID="Table1" runat="server" Visible="true" HorizontalAlign="Center" />
            <p></p><hr /><p></p>
        <h3>Statistical Analysis</h3> <p>
        <asp:Table ID="Table4" runat="server" Visible="true" HorizontalAlign="Center" />
        </p><p></p><hr /><p></p>
        <h3>Analytic Model Options</h3> 
        <p></p>
        <!doctype html>
        <html lang="en">
        <head>
            <meta charset="utf-8">
            <title>accordion</title>
            <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
            <script src="//code.jquery.com/jquery-1.10.2.js"></script>
            <script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
            <script>
                $(function () {
                    $("#accordion").accordion({
                        collapsible: true
                    });
                });
            </script>
        </head>
        <body>
 
        <div id="accordion">
            <% if(Table2.Rows[1].Cells[0].Text != "-") %>
            <% { %>
                  <h3>Please Resolve Missing Values</h3>
                  <div>
                    <p>It's critical that data submitted for statistical analysis contains only the supported NULL value indicators. 
                        Leaving a row element blank could lead to errors, inaccurate characterization of the data, and unexpected results. </p>
                  </div>          
            <% } %>
            <% if(true == true) %>
            <% { %>
                  <h3>Please Select Variable Dependency</h3>
                  <div>
                    <p>Without selecting at least 1 independent variable and 1 dependent variable, you cannot perform regression analysis.</p>
                  </div>          
            <% } %>
         </div>
 
        <script>
            $("#accordion").accordion();
        </script>
        </body> 
        </html>
        </p>
    </div>
</asp:Content>

