using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using eBay.Service.Core.Soap;
using eBay.Service.Core.Sdk;
using eBay.Service.Call;
using eBay.Service.Util;
using eBay.Service.EPS;

namespace SoapLibraryDemo
{
    public partial class FrmUploadPictures : Form
    {

        public ApiContext Context;

        public List<string> fileList = new List<string>();

        public FrmUploadPictures()
        {
            InitializeComponent();
        }

        //find an image
        private void browseButton_Click(object sender, EventArgs e)
        {
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog.FileNames != null)
                {
                    for (int i = 0; i < openFileDialog.FileNames.Length; i++)
                    {
                        addImage(openFileDialog.FileNames[i]);
                    }
                }
                else
                    addImage(openFileDialog.FileName);

                this.updateListView();
            }

        }

        private void addImage(string imageToLoad)
        {
            if (imageToLoad != "")
            {
                this.fileList.Add(imageToLoad);
            }
        }

        private void updateListView()
        {
            try
            {
                imageList.Images.Clear();
                listView.Items.Clear();
                for (int i = 0; i < this.fileList.Count; i++)
                {
                    string file = this.fileList[i];
                    imageList.Images.Add(ResizeImage(Image.FromFile(file), 100, 100));
                    ListViewItem lvi = new ListViewItem(Path.GetFileName(file));
                    lvi.ImageIndex = i;
                    this.listView.Items.Add(lvi);
                }
            }
            catch (Exception ex)
            {
                this.fileList.Clear();
                imageList.Images.Clear();
                listView.Items.Clear();
                MessageBox.Show(ex.Message);
            }
        }

        private static System.Drawing.Bitmap ResizeImage(System.Drawing.Image image,
                                                int width, int height)
        {
            //a holder for the result 
            Bitmap result = new Bitmap(width, height);

            //use a graphics object to draw the resized image into the bitmap 
            using (Graphics graphics = Graphics.FromImage(result))
            {
                //set the resize quality modes to high quality 
                graphics.CompositingQuality =
                    System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                graphics.InterpolationMode =
                    System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode =
                    System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                //draw the image into the target bitmap 
                graphics.DrawImage(image, 0, 0, result.Width, result.Height);
            }
            //return the resulting bitmap 
            return result;
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            if (this.listView.SelectedItems.Count > 0)
            {
                StringCollection sc = new StringCollection();
                foreach (int index in this.listView.SelectedIndices)
                {
                    sc.Add(this.fileList[index]);
                }
                foreach (string file in sc)
                {
                    this.fileList.Remove(file);
                }
                this.updateListView();
            }
        }

        private void uploadButton_Click(object sender, EventArgs e)
        {
            try
            {
                eBayPictureService eps = new eBayPictureService(Context);
                string result = "";
                int size = this.fileList.Count;
                //upload pictures one by one
                for (int i = 0; i < size; i++)
                {
                    string file = this.fileList[i];
                    UploadSiteHostedPicturesRequestType request = new UploadSiteHostedPicturesRequestType();
                    if (this.checkBox.Checked)
                    {
                        request.PictureWatermark = new PictureWatermarkCodeTypeCollection();
                        request.PictureWatermark.Add(PictureWatermarkCodeType.User);
                    }
                    if (this.extDaysTextBox.Text != String.Empty)
                    {
                        int extDays = int.Parse(this.extDaysTextBox.Text);
                        request.ExtensionInDays = extDays;
                    }
                    UploadSiteHostedPicturesResponseType response = eps.UpLoadSiteHostedPicture(request, file);
                    result += Path.GetFileName(file) + " : " + response.SiteHostedPictureDetails.FullURL + "\r\n";
                }
                this.textBox.Text = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

        }
    }
}