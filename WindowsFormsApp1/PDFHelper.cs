using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Configuration;

namespace MarketingPDFMerge
{
    class PDFHelper
    {
        public static string pdfPassword = ConfigurationManager.AppSettings["PDF_Password"];
        public static string inputPassword = "20Years";

        /**
         * PDF Settings
         **/
        public void saveSettings(string filePath)
        {            
            Document doc = new Document();

            PdfReader reader = new PdfReader(filePath);

            using (MemoryStream docStream = new MemoryStream())
            {

                PdfWriter writer = PdfWriter.GetInstance(doc, docStream);

                writer.ViewerPreferences = PdfWriter.FitWindow;

                doc.Open();

                PdfContentByte pdfByte = writer.DirectContent;

                float zoomNumber = 75;
                int pageNumberToOpenTo = 1;

                PdfDestination magnify = new PdfDestination(PdfDestination.XYZ, -1, -1, zoomNumber / 100);

                PdfAction zoom = PdfAction.GotoLocalPage(pageNumberToOpenTo, magnify, writer);

                reader.Close();

                doc.Close();

                File.WriteAllBytes(filePath, docStream.ToArray());

            }

        }

        /**
         * PDF Page Selector To Trim PDF
         **/
        public void SelectPages(string inputPdf, string pageSelection, string outputPdf)
        {
            using (PdfReader reader = new PdfReader(inputPdf))
            {
                reader.SelectPages(pageSelection);

                using (PdfStamper stamper = new PdfStamper(reader, File.Create(outputPdf)))
                {
                    stamper.Close();
                }
            }
        }

        /**
         * Merge Multiple PDFs
         **/
        public bool MergePDFs(IEnumerable<string> fileNames, string targetPdf)
        {
            bool merged = true;
            using (FileStream stream = new FileStream(targetPdf, FileMode.Create))
            {
                iTextSharp.text.Document document = new iTextSharp.text.Document();
                PdfCopy pdf = new PdfCopy(document, stream);
                PdfReader reader = null;
                try
                {
                    document.Open();
                    foreach (string file in fileNames)
                    {
                        reader = new PdfReader(file);
                        pdf.AddDocument(reader);
                        reader.Close();
                    }
                }
                catch (Exception)
                {
                    merged = false;
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
                finally
                {
                    if (document != null)
                    {
                        document.Close();
                    }
                }
            }
            return merged;
        }

        /**
         * Decrypt File
         **/
        private void DecryptFile(string inputFile, string outputFile)
        {

            try
            {
                PdfReader reader = new PdfReader(inputFile, new System.Text.ASCIIEncoding().GetBytes(inputPassword));

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    PdfStamper stamper = new PdfStamper(reader, memoryStream);
                    stamper.Close();
                    reader.Close();
                    File.WriteAllBytes(outputFile, memoryStream.ToArray());
                }

            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        /**
         * ENcrypt PDF
         * */
        public void encryptPDF(string unsecured, string secured)
        {
            string InputFile = unsecured;
            string OutputFile = secured;

            using (Stream input = new FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (Stream output = new FileStream(OutputFile, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfReader reader = new PdfReader(input);

                    PdfEncryptor.Encrypt(reader, output, true, null, pdfPassword, PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING);
                }
            }
        }

    }
}
