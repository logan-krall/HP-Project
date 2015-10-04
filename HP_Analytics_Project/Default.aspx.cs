﻿using System;
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
                string fn, save = String.Empty;

                string path = Server.MapPath("~//Uploads/");

                //string path = HttpContext.Current.Server.MapPath("~//Uploads/");

                //string virtualPathToDirectory = "~/Uploads/";
                //string path = Server.MapPath(virtualPathToDirectory);

                if (FileUpload1.HasFile)
                {
                    string[] allowedExtensions = { ".xls", ".xlsx", ".csv" };
                    string extension = System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();
                    if (allowedExtensions.Contains(extension))
                    {
                        int filesize = FileUpload1.PostedFile.ContentLength;
                        if (filesize < 25000000)
                        {
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
                        fn = System.IO.Path.GetFileName(FileUpload1.PostedFile.FileName);
                        save = Server.MapPath(fn);
                        FileUpload1.PostedFile.SaveAs(save);
                        Session["name"] = save;
                        //FileUpload1.PostedFile.SaveAs(path + FileUpload1.FileName);

                         UploadStatusLabel.Text = "File uploaded successfully.";
                    }
                        catch (Exception err)
                    {
                        UploadStatusLabel.Text = "File could not be uploaded because " + err + " exception caught.";
                    }

                    //if (File.Exists(path + FileUpload1.FileName))
                    if (File.Exists(save))
                    {
                        Server.Transfer("Upload.aspx", true);
                    }
                }
            }
        }
    }
}