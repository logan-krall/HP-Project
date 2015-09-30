using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

namespace HP_Analytics_Project
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void UploadButton_Click(object sender, EventArgs e)
        {
            if (IsPostBack)
            {
                Boolean fileOK = false;
                String path = Server.MapPath("~/Uploads/");
                if (FileUpload1.HasFile)
                {
                    string[] allowedExtensions = { ".xls", ".xlsx", ".csv" };
                    string extension = System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();
                    if (allowedExtensions.Contains(extension))
                    {
                        int filesize = FileUpload1.PostedFile.ContentLength;
                        if (filesize < 25000000)
                        {
                            Session["name"] = FileUpload1.FileName;
                            fileOK = true;
                        }
                        else { UploadStatusLabel.Text = "Your file was not uploaded because it is too large."; }
                    }
                    else { UploadStatusLabel.Text = "Your file was not uploaded because it was not a supported file type"; }
                }
                else { UploadStatusLabel.Text = "You must select a file to upload."; }
                if (fileOK)
                {
                    if (File.Exists(path + FileUpload1.FileName))
                    {
                        File.Delete(path + FileUpload1.FileName);
                    }
                    try
                    {
                        FileUpload1.PostedFile.SaveAs(path + FileUpload1.FileName);
                        UploadStatusLabel.Text = "File uploaded successfully.";
                    }
                    catch
                    {
                        UploadStatusLabel.Text = "File could not be uploaded.";
                    }
                    if (File.Exists(path + FileUpload1.FileName))
                    {
                        Server.Transfer("Upload.aspx", true);
                    }
                }
            }
        }
    }
}