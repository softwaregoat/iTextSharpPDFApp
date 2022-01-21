using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Font = iTextSharp.text.Font;
using Rectangle = iTextSharp.text.Rectangle;
using System.Configuration;
using System.Security.AccessControl;

namespace MarketingPDFMerge
{

    public partial class Form1 : Form
    {
        private static string PresentationsFolder = ConfigurationManager.AppSettings["PRESENTATIONS_FOLDER"];


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var date = " ";
            if (cmbYear.Text == "" || cmbMonth.Text == "")
            {
                date = " ";
            }
            else
            {
                date = $"{cmbMonth.Text} {cmbDay.Text}, {cmbYear.Text}";
            }
            //var date = dateTimePicker1.Value.ToString("MMMM dd, yyyy");
            var now = DateTime.Now.ToString("yyyyMMdd-hhmmss");

            var logo_path = PresentationsFolder + @"Images\logo.png";
            var font_path = PresentationsFolder + @"Font\COP32BC 2021-11-11 19_25_33.TTF";
            var date_font_path = PresentationsFolder + @"Font\GoudySAMG 2021-11-11 19_25_38.ttf";


            BaseFont bf = BaseFont.CreateFont(font_path, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            //Font f = new Font(bf, 12);
            Font fClient = new Font(bf, 18);
            fClient.SetColor(22, 53, 116);

            Font fPresenter = new Font(bf, 16);
            fPresenter.SetColor(22, 53, 116);

            Font fTitle = new Font(bf, 14);
            fTitle.SetColor(22, 53, 116);


            BaseFont date_bf = BaseFont.CreateFont(date_font_path, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font date_f = new Font(date_bf, 12);

            Font footer_f = new Font(bf, 8);


            var client_name = txtClientName.Text;

            var presenter1 = cmbPresenter1.Text;
            var presenter1s = presenter1.Split(';');

            if (presenter1s.Length != 2)
            {
                MessageBox.Show("Please select Presenter 1");
                return;
            }

            var present_name1 = presenter1s[0];
            var present_title1 = presenter1s[1];

            var presenter2 = cmbPresenter2.Text;
            var presenter2s = presenter2.Split(';');

            var present_name2 = " ";
            var present_title2 = " ";

            if (presenter2s.Length == 2)
            {
                present_name2 = presenter2s[0];
                present_title2 = presenter2s[1];
            }

            var presentation = cmbPresentation.Text;
            if (presentation == "")
            {
                MessageBox.Show("Please select Presentation");
                return;
            }


            var presentation_template = presentation + ".pdf";

            var presentation_file = PresentationsFolder + presentation_template;

            var orient = "Portrait";
            if (File.Exists(presentation_file))
            {
                byte[] bytes = System.IO.File.ReadAllBytes(presentation_file);
                using (MemoryStream ms = new MemoryStream())
                {
                    PdfReader reader = new PdfReader(bytes);
                    Rectangle currentPageRectangle = reader.GetPageSizeWithRotation(1);
                    if (currentPageRectangle.Width > currentPageRectangle.Height)
                    {
                        //page is landscape
                        orient = "Landscape";
                    }
                    else
                    {
                        //page is portrait
                        orient = "Portrait";
                    }
                }
            }

            Rectangle rec = new Rectangle(PageSize.LETTER);


            //var dest_path = @"P:\AllUsers\" +Environment.UserName + @"\_SilvercrestRpts";
            var dest_path = @"E:\_SilvercrestRpts";

            if(!System.IO.Directory.Exists(dest_path))
            {
                System.IO.Directory.CreateDirectory(dest_path);
            }

            dest_path = dest_path + @"\CustomMarketingMaterials";
            if (!System.IO.Directory.Exists(dest_path))
            {
                System.IO.Directory.CreateDirectory(dest_path);
            }

            dest_path = dest_path  + @"\Marketing_Presentation_" + now + ".pdf";
            if (orient == "Landscape")
            {
                rec = new Rectangle(PageSize.LETTER.Rotate());
            }

            string timestamp = DateTime.Now.ToString("hhmmss") + ".pdf";
            string tmpPDF = PresentationsFolder + @"Temp\TMP_Marketing_Presentation_" + timestamp;
            string tmpPDF_Cover = PresentationsFolder + @"Temp\TMP_Marketing_Presentation_Cover_" + timestamp;
            string tmpPDF_Final = PresentationsFolder + @"Temp\TMP_Marketing_Presentation_Final_" + timestamp;

            /**Generating Cover Page***/
            using (FileStream fs = new FileStream(tmpPDF_Cover, FileMode.Create, FileAccess.Write, FileShare.None))
            using (Document doc = new Document(rec))
            using (PdfWriter writer = PdfWriter.GetInstance(doc, fs))
            {
                doc.Open();

                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(logo_path);
                logo.Alignment = Element.ALIGN_CENTER;
                logo.SetAbsolutePosition(210, 480);
                if (orient == "Landscape") {
                    logo.SetAbsolutePosition(300, 350);
                }
                logo.ScalePercent(20);
                doc.Add(logo);

                Paragraph lineSeparator = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(1F, 10.0F, BaseColor.GRAY, Element.ALIGN_CENTER, 1)));
                if (orient == "Landscape")
                {
                    lineSeparator.SpacingBefore = 230;
                }
                else
                {
                    lineSeparator.SpacingBefore = 310;
                    lineSeparator.SpacingAfter = 20;
                }
                doc.Add(lineSeparator);

                //ColumnText ct = new ColumnText(writer.DirectContent);
                //ct.SetSimpleColumn(new Rectangle(doc.PageSize.Width/2 - 150, 400, doc.PageSize.Width / 2 + 150, 450));
                //ct.AddElement(new Paragraph("This could be a very long sentence that needs to be wrapped"));
                //ct.AddElement(lineSeparator);
                //ct.Go();

                //ColumnText.ShowTextAligned(writer.DirectContent, Element.ALIGN_CENTER, new Phrase(lineSeparator), 150, 460, 0);

                Paragraph para = new Paragraph("PRESENTATION TO", fClient);
                para.Alignment = Element.ALIGN_CENTER;
                para.SpacingBefore = 15;
                doc.Add(para);

                para = new Paragraph(client_name.ToUpper(), fClient);
                para.Alignment = Element.ALIGN_CENTER;
                doc.Add(para);


                para = new Paragraph(date, date_f);
                para.Alignment = Element.ALIGN_CENTER;
                para.SpacingBefore = 20;
                doc.Add(para);

                para = new Paragraph(present_name1.ToUpper(), fPresenter);
                para.Alignment = Element.ALIGN_CENTER;
                para.SpacingBefore = 20;
                doc.Add(para);

                para = new Paragraph(present_title1.ToUpper(), fTitle);
                para.Alignment = Element.ALIGN_CENTER;
                para.SetLeading(1.0f, 1.0f);
                doc.Add(para);

                para = new Paragraph(present_name2.ToUpper(), fPresenter);
                para.Alignment = Element.ALIGN_CENTER;
                para.SpacingBefore = 20;
                doc.Add(para);

                para = new Paragraph(present_title2.ToUpper(), fTitle);
                para.Alignment = Element.ALIGN_CENTER;
                para.SetLeading(1.0f, 1.0f);
                doc.Add(para);


                if(orient == "Portrait")
                {
                    para = new Paragraph("SILVERCRESTASSET MANAGEMENT GROUP LLC", footer_f);
                    if(present_name2.Length == 0)
                    {
                        para.SpacingBefore = 120;
                    }
                    else
                    {
                        para.SpacingBefore = 90;
                    }
                    //if (orient == "Landscape")
                    //{
                    //    para.SpacingBefore = 50;
                    //}
                    para.Alignment = Element.ALIGN_CENTER;
                    doc.Add(para);


                    para = new Paragraph("1330  AVENUE OF THE AMERICAS, NEW YORK, NEW YORK 10019  • (212) 649-0600", footer_f);
                    para.Alignment = Element.ALIGN_CENTER;
                    doc.Add(para);


                    para = new Paragraph("WWW.SILVERCRESTGROUP.COM", footer_f);
                    para.Alignment = Element.ALIGN_CENTER;
                    doc.Add(para);

                }

                float zoomNumber = 75.0f;
                int pageNumberToOpenTo = 1;

                PdfDestination magnify = new PdfDestination(PdfDestination.XYZ, -1, -1, zoomNumber / 100);

                writer.ViewerPreferences = PdfWriter.FitWindow;

                PdfAction zoom = PdfAction.GotoLocalPage(pageNumberToOpenTo, magnify, writer);

                doc.Close();
            }


            /*****PDF Construction******/
            PDFHelper obj = new PDFHelper();

            try
            {
                // Open the file
                PdfReader inputDocument = new PdfReader(presentation_file); //Presentation Template Files
                int pageCount = inputDocument.NumberOfPages;

                //Stripping out Page 1
                obj.SelectPages(presentation_file, "2-" + pageCount.ToString(), tmpPDF);

                //Merging PDFs
                List<string> pdfs = new List<string>();
                pdfs.Add(tmpPDF_Cover);
                pdfs.Add(tmpPDF);

                obj.MergePDFs(pdfs, tmpPDF_Final);
//                obj.saveSettings(tmpPDF_Final);
                obj.encryptPDF(tmpPDF_Final, dest_path);

                File.Delete(tmpPDF_Cover);

                File.Delete(tmpPDF);

                File.Delete(tmpPDF_Final);

                string outputMessage = string.Format("Your presentation is now available here: {0}{0}{1}", Environment.NewLine, dest_path);

                MessageBox.Show(outputMessage);

            }
            catch (Exception x)
            {
                if (File.Exists(tmpPDF))
                {
                    File.Delete(tmpPDF);
                }

                if (File.Exists(tmpPDF_Final))
                {
                    File.Delete(tmpPDF_Final);
                }

                if (File.Exists(tmpPDF_Cover))
                {
                    File.Delete(tmpPDF_Cover);
                }

            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var presenter_path = PresentationsFolder + "Presenters.txt";
            var presentation_path = PresentationsFolder + "Presentations.txt";

            foreach (string line in System.IO.File.ReadLines(presenter_path))
            {
                cmbPresenter1.Items.Add(line);
                cmbPresenter2.Items.Add(line);
            }

            foreach (string line in System.IO.File.ReadLines(presentation_path))
            {
                cmbPresentation.Items.Add(line);
            }

            for (int i = 1; i < 32; i++)
            {
                cmbDay.Items.Add(i.ToString("00"));
            }
            string[] monthNames = new string[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            cmbMonth.Items.AddRange(monthNames);

            for (int i = 2021; i < 2041; i++)
            {
                cmbYear.Items.Add(i.ToString());
            }

            var now = DateTime.Now;
            cmbDay.Text = now.Day.ToString("00");
            cmbMonth.SelectedIndex = now.Month - 1;
            cmbYear.Text = now.Year.ToString();
        }
    }
}
