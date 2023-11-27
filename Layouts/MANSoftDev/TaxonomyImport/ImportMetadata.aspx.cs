using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Xml.Linq;
using System.IO;

namespace TaxonomyImport.Layouts.TaxonomyImport
{
    public partial class ImportMetadata : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// Import button click handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void OnImport(object sender, EventArgs e)
        {
            if(FileUpload.HasFile)
            {
                try
                {
                    TextReader reader = new StreamReader(FileUpload.PostedFile.InputStream);
                    MANSoftDev.SharePoint.Utilities.Taxonomy.TaxonomyHelper.ImportTerms(SPContext.Current.Site, reader);

                    msg.InnerText = "Import complete";
                    msg.Visible = true;
                }
                catch(Exception ex)
                {
                    msg.InnerText = ex.Message;
                    msg.Visible = true;
                }
            }
            else
            {
                msg.InnerText = "Please select a file.";
                msg.Visible = true;
            }
        }
    }
}
