<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="HP_Analytics_Project._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <img runat="server" src="../Images/abacus_white1.png" style="width:auto;height:70px;">
            </hgroup>
            <span style="color:#fff">This analytics application was written to prescreen data sets for statistical analysis and to help educate and explore different statistical model options in the pursuit of generating actionable data from big data sets. </span>
        </div>
    </section>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    
    <h3>&nbsp;Please select a file to upload:</h3>
    &nbsp;&nbsp;File must be less than 250mb and be of the .csv, .xls, or .xlsx extension types. <br/><br/><br/>
    <center>
    &nbsp;<asp:FileUpload id="FileUpload1" runat="server"> </asp:FileUpload>
    <asp:Button ID="UploadButton" Text="Upload File" OnClick="UploadButton_Click" runat="server" Font-Size="Small "/> 
    </center>   
    </br></br>
    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="UploadStatusLabel" runat="server"></asp:Label>


</asp:Content>
