<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Upload.aspx.cs" Inherits="HP_Analytics_Project.Images.WebForm1" Async="true" %>



<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1>Abacus</h1>
            </hgroup>
            <span style="color:#fff">This analytics application was written to prescreen data sets for statistical analysis and to help educate and explore different statistical model options in the pursuit of generating actionable data from big data sets. </span>
        </div>
    </section>
</asp:Content>



<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server" >
    
    <div style="text-align:center">
        <head>                       
            <link rel="stylesheet" type="text/css" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
            <script type="text/javascript" src="//code.jquery.com/jquery-1.10.2.js"></script>
            <script type="text/javascript" src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
            
            <script type="text/javascript">
                function BindEvents() {
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
                            active: false,

                        });
                    });
                    $(function () {
                        $("#accordion4").accordion({
                            collapsible: true,
                            heightStyle: "content",
                            clearStyle: true,
                        });
                    });
                }
            </script>

            
        </head>
        <body>            
        
        <% if ((bool)ViewState["missing"] == true) %>
        <% { %>
            <h3 style="color:red">Missing Row Values</h3> 
            <asp:Table ID="Table2" runat="server" Visible="true" HorizontalAlign="Center" BackColor="White" BorderStyle="Solid" BorderWidth="1"/>
            <p/><p/>
            ** Please resolve all instances of rows with missing values prior to submission for statistical analysis. **
            <p /> <p />
            
            <p />
            Click here to save a new copy of the spreadsheet in xls format with the missing values replaced by average or placeholder values. 
            <p />
            <asp:Button OnClick="saveButton_Click" id="saveButton" Text="Save New" runat="server" />
            <p />
            <asp:Label ID="UploadStatusLabel" runat="server"/>
            <p/><hr /><p/>
        <% } %>    

        <%-- Ajax for changing without postback --%>
            <asp:UpdatePanel ID="panel1" runat="server">  
                <ContentTemplate>

                <%-- Rebind events --%>
                <script type="text/javascript">
                    Sys.Application.add_load(BindEvents);
                </script>

                <h3>Data Characteristics</h3>
                <p></p>
                <div id="accordion1">
                <h3>Numeric Statistics 
                <% if ((int)ViewState["num_numeric"] > 0) { %>
                    ( <%= (int)ViewState["num_numeric"] %> 
                    <% if ((int)ViewState["num_numeric"] == 1) { %>
                    Variable )
                    <% } %>
                    <% else { %>
                    Variables )
                    <% } %>
                <% } %>
                <% else { %>
                    ( Empty )
                <% } %>
                </h3>
                <div>
                <asp:Table ID="Table1" runat="server" Visible="true" HorizontalAlign="Center" />
                <p></p>
                </div>
                </div>
                <div id="accordion3">
                    <h3>Nominal Statistics 
                        <% if ((int)ViewState["num_nominal"] > 0) { %>
                            ( <%= (int)ViewState["num_nominal"] %> 
                        <% if ((int)ViewState["num_nominal"] == 1) { %>
                        Variable )
                        <% } %>
                        <% else { %>
                        Variables )
                        <% } %>
                    <% } %>
                    <% else { %>
                        ( Empty )
                    <% } %>

                    </h3>
                    <div>
                    <asp:Table ID="Table3" runat="server" Visible="false" HorizontalAlign="Center" />
                    <p></p>
                    </div>
                </div>
                <p /><p/><hr /><p/>
                <h3>Correlation Matrix</h3>

                <%-- Correlational Matrix Table --%>
                <asp:Table ID="CorrTable" runat="server" Visible="true" HorizontalAlign="Center" BackColor="White" BorderStyle="Solid" BorderWidth="1"/>
                <p /><hr /><p />
                <h3>Analytic Model Options</h3> 
                <p/> 
                <div id="accordion4">
                    <% if(Table2.Rows.Count > 0 && Table2.Rows[1].Cells[0].Text != "-") %>
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
                          <h3>Simple Regression</h3>
                          <div>
                          <img src="../Images/lin-reg.png" style="width:400px;height:300px;">
                            <p>
                                Simple regression analysis is used to describe data and help explain the relationship between a single independent variable and 
                                a dependent variable but produces results that are typically less realistic than multiple regression analysis because of the 
                                limited number of independent, or explanatory, variables. The three most common applications of regression analysis includes 
                                being able to perform causal analysis to identify trends, predictive analysis on individual cases, and forecasting overall trends. 
                            </p><p>
                                An example of simple regression being utilized in a business setting could provide understanding on how sensitive sales are to 
                                changes in total advertising expenditures. Such an analysis would provide the ability to make predictions about increases or 
                                decreases in sales volume based on a specific scenario where changes to advertising are made. 
                            </p>
                          </div>
            
                           <h3>Ridge Regression</h3>
                          <div>
                          <img src="../Images/ridge_regression.png" style="width:400px;height:400px;">
                            <p>
                                Ridge regression analysis can be performed between multiple independent, or explanatory, variables and a dependent variable, 
                                similarly to multiple regression, but is typically used when your explanatory variables are highly correlated with each other. 
                                Ridge regression may also be used when your dependent variable only represents a small number of values, which causes 
                                Least Squares regression analysis, an alternative to ridge regression, to produce inaccurate predictions. 
                            </p><p>
                                For regression analysis, you want a high degree of correlation between your explanatory and dependent variables so that 
                                accurate predictions can be made, but when you have highly correlated explanatory variables the predictions may be less 
                                accurate. The degree of correlation between variables can be calculated and displayed in a correlational matrix with a 
                                percentage value representing their relationship at the intersection of their sets in the table. 
                            </p><p>
                                An example of ridge regression being utilized in a business setting could provide understanding on how sensitive sales are 
                                to changes in advertising dollars spent, and various additional influences. With these results, forecasts for future 
                                demand can be generated and used to predict changes in sales that could be either beneficial or detrimental to profits. 
                            </p>
                          </div>   
                
                          <h3>Lasso Regression</h3>
                          <div>
                          <img src="../Images/ordinary_lasso.png" style="width:400px;height:400px;">
                            <p>
                                Least Absolute Shrinkage of Selection Operator, or Lasso for short, can be performed between multiple independent, or explanatory, 
                                variables and a dependent variable, similarly to multiple and ridge regression analysis. In contrast, it allows for both highly 
                                correlated explanatory variables, which is a problem for multiple regression, and the ability to shrink the influence of certain 
                                variables down to zero, which cannot be done in ridge regression. 
                            </p><p>
                                For regression analysis, you want a high degree of correlation between your explanatory and dependent variables so that accurate 
                                predictions can be made, but when you have highly correlated explanatory variables the predictions may be less accurate. The degree 
                                of correlation between variables can be calculated and displayed in a correlational matrix with a percentage value representing their 
                                relationship at the intersection of their sets in the table. 
                            </p><p>
                                An example of lasso regression being utilized in a business setting could provide understanding on how sensitive sales are to changes 
                                in advertising dollars spent, and various additional influences. With these results, forecasts for future demand can be generated and 
                                used to predict changes in sales that could be either beneficial or detrimental to profits. 
                            </p>
                          </div>  
            
                          <h3>Bayesian Regression</h3>
                          <div>
                  
                            <p>
                                Bayesian regression analysis can be performed between multiple independent, or explanatory, variables and a dependent variable, 
                                similarly to multiple and ridge regression analysis. In contrast to ridge regression, bayesian regression allows for the compensation 
                                of overfitting when a model has too many parameters relative to the number of observations by relaxing the assumption that the errors 
                                must have a normal distribution. 
                            </p><p>
                                An example of bayesian regression being utilized in a business setting could provide understanding on how sensitive sales are to changes
                                in advertising dollars spent, and various additional influences. With these results, forecasts for future demand can be generated and 
                                used to predict changes in sales that could be either beneficial or detrimental to profits. 
                            </p>
                          </div>       
                    <% } %>
                    <% if (Multi_Reg_Check() == true) %>
                    <% { %>
                           <h3>Multiple Regression</h3>
                          <div>
                            <img src="../Images/mult-reg.png" style="width:400px;height:300px;">
                            <p>
                                Multiple regression analysis can be performed between multiple independent variables and a single dependent variable but produces 
                                results that are typically more realistic than simple regression analysis because of the increased number of independent, or 
                                explanatory, variables. The three most common applications of regression analysis includes being able to perform causal analysis 
                                to identify trends, predictive analysis on individual cases, and forecasting overall trends. 
                            </p>
                            <p>
                                An example of multiple regression being utilized in a business setting could provide understanding on how sensitive sales are to 
                                changes in advertising dollars spent and various additional influences. With these results, forecasts for future demand 
                                can be generated and used to predict changes in sales that could be either beneficial or detrimental to profits. 
                            </p>
                          </div>    
                    <% } %>
                    <% if (Logit_Reg_Check() == true) %>
                    <% { %>
                        <h3>Logit Regression</h3>
                        <div>
                            <img src="../Images/logit-reg.png" style="width:400px;height:300px;">
                            <p>
                                Logistic regression analysis can be performed between a single independent variable and a dependent variable, similarly to simple 
                                regression, but requires a categorical dependent variable rather than continuous. While multinomial logistic regression can be 
                                performed on a categorical variable with a finite range of values, logistic regression is typically used on dichotomous variables 
                                that only represent two possible outcomes. 
                            </p><p>
                                An example of logistic regression on a dichotomous categorical dependent variable in a business setting could be performing an 
                                analysis on customer retention using length of the customer's account history. Customer retention is a two outcome variable that 
                                represents making another sale or not and using this analysis we can make predictions about the probability of future sales. 

                            </p>
                        </div>
                    <% } %>
                 </div>
 
                <script>
                    $("#accordion").accordion();
                </script>
            
                </ContentTemplate>
            </asp:UpdatePanel>

        </body> 
        </p>
    </div>
</asp:Content>

