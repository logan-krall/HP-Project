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



<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server" >
    
    <div style="text-align:center">
        <%-- <!doctype html> --%>
        <%-- <html lang="en"> --%>
        <head>
            <%-- <meta charset="utf-8"> --%>
            <title>accordion</title>
            <link rel="stylesheet" type="text/css" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
            <script type="text/javascript" src="//code.jquery.com/jquery-1.10.2.js"></script>
            <script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
            
            <script type="text/javascript">
                $(function () {
                    $("#accordion1").accordion({
                        collapsible: true,
                        heightStyle: "content",
                        clearStyle: true,
                    });
                });
                $(function () {
                    $("#accordion2").accordion({
                        collapsible: true,
                        heightStyle: "content",
                        clearStyle: true,
                    });
                });
                $(function () {
                    $("#accordion3").accordion({
                        collapsible: true,
                        heightStyle: "content",
                        clearStyle: true,
                    });
                });
                $(function () {
                    $("#accordion4").accordion({
                        collapsible: true,
                        heightStyle: "content",
                        clearStyle: true,
                    });
                });
        </script>

            
        </head>
        <body>            

         

        <% if ((bool)ViewState["missing"] == true) %>
        <% { %>
            <h3>Missing Row Values</h3> 
            <asp:Table ID="Table2" runat="server" Visible="true" HorizontalAlign="Center" BackColor="White" BorderStyle="Solid" BorderWidth="1"/>
            <p></p><hr /><p></p>
        <% } %>    

        <h3>Data Characteristics</h3>
        <p></p>
        <div id="accordion2">
            <h3>Numeric Statistics</h3>
            <div>
            <asp:Table ID="Table1" runat="server" Visible="true" HorizontalAlign="Center" />
            <p></p>
            </div>
        
            <h3>Nominal Statistics</h3>
            <div>
            <asp:Table ID="Table3" runat="server" Visible="true" HorizontalAlign="Center" />
            <p></p>
            </div>
        </div>
        <p></p><hr /><p></p>

        <% if (Depend_Check() == true) %>
        <% { %>
            <h3>Correlation Matrix</h3>
            <asp:Table ID="Table4" runat="server" Visible="true" HorizontalAlign="Center" BackColor="White" BorderStyle="Solid" BorderWidth="1" />
            <p></p><hr /><p></p><p></p>
        <% } %>

        <h3>Analytic Model Options</h3> 
        <p></p> 
        <div id="accordion4">
            <% if(Table2.Rows[1].Cells[0].Text != "-") %>
            <% { %>
                  <h3>Please Resolve Missing Values</h3>
                  <div>
                    <p>
                        It's critical that data submitted for statistical analysis contains only the supported NULL value indicators. 
                        Leaving a row element blank could lead to errors, inaccurate characterization of the data, and unexpected results. 
                    </p>
                  </div>          
            <% } %>
            <% if(Depend_Check() == false) %>
            <% { %>
                  <h3>Please Select Variable Dependency</h3>
                  <div>
                    <p>Without selecting at least 1 independent variable and 1 dependent variable, you cannot perform regression analysis.</p>
                  </div>          
            <% } %>
            <% else %>
            <% { %>
                  <h3>Linear Regression</h3>
                  <div>
                  <img src="../Images/lin-reg.png" style="width:400px;height:300px;">
                    <p>
                        Linear regression is the most basic and commonly used predictive analysis.  
                        Regression estimates are used to describe data and to explain the relationship 
                        between one dependent variable and one or more independent variables. 
                        There are 3 major uses for regression analysis – (1) causal analysis, (2) forecasting an effect, (3) trend forecasting.
                    </p>
                    <p>
                        However linear regression analysis consists of more than just fitting a linear line through a cloud of data points.  
                         It consists of 3 stages – (1) analyzing the correlation and directionality of the data, (2) estimating the model, i.e., 
                         fitting the line, and (3) evaluating the validity and usefulness of the model.
                    </p>
                  </div>          
            <% } %>
            <% if (Multi_Reg_Check() == true) %>
            <% { %>
                   <h3>Multiple Linear Regression</h3>
                  <div>
                    <img src="../Images/mult-reg.png" style="width:400px;height:300px;">
                    <p>
                        Multiple Linear regression is the most basic and commonly used predictive analysis.  
                        Regression estimates are used to describe data and to explain the relationship 
                        between one dependent variable and two or more independent variables. 
                        There are 3 major uses for regression analysis – (1) causal analysis, (2) forecasting an effect, (3) trend forecasting.
                    </p>
                    <p>
                        However linear regression analysis consists of more than just fitting a linear line through a cloud of data points.  
                         It consists of 3 stages – (1) analyzing the correlation and directionality of the data, (2) estimating the model, i.e., 
                         fitting the line, and (3) evaluating the validity and usefulness of the model.
                    </p>
                  </div>    
            <% } %>
            <% if (Logit_Reg_Check() == true) %>
            <% { %>
                <h3>Logit Regression</h3>
                <div>
                    <img src="../Images/logit-reg.png" style="width:400px;height:300px;">
                    <p>
                        Logistic regression, also called a logit model, is used to model dichotomous outcome variables. 
                        In the logit model the log odds of the outcome is modeled as a linear combination of the predictor variables.
                    </p>
                </div>
            <% } %>
         </div>
 
        <script>
            $("#accordion").accordion();
        </script>
        </body> 
        <%--</html> --%>
        </p>
    </div>
</asp:Content>

