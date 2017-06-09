using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace SECExtract {
    public class PdfExtractor {
        public static int GetPageCount(string sourcePdfPath) {
            int pageCount = 0;
            try {
                var reader = new PdfReader(sourcePdfPath);
                pageCount = reader.NumberOfPages;
                reader.Close();
            }
            catch (Exception) {
                throw;
            }
            return pageCount;
        }

        public static string ReadPdfFile(string fileName, int startPage = 1, int endPage = -1) {
            StringBuilder text = new StringBuilder();

            if (File.Exists(fileName)) {
                PdfReader pdfReader = new PdfReader(fileName);

                if (startPage < 1 || startPage > endPage)
                    startPage = 1;

                if (endPage < 0 || endPage < startPage)
                    endPage = pdfReader.NumberOfPages;

                for (int page = startPage; page <= endPage; page++) {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        public static bool ExtractPages(string sourcePdfPath, string outputPdfPath, int startPage, int endPage) {
            bool success = false;
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try {
                reader = new PdfReader(sourcePdfPath);

                sourceDocument = new Document(reader.GetPageSizeWithRotation(startPage));

                pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                sourceDocument.Open();

                for (int i = startPage; i <= endPage; i++) {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
                success = true;
            }
            catch {
                throw;
            }

            return success;
        }

        public static List<string> SplitPDF(string pdfFromDoc, string dumpPath) {
            var output = new List<string>();

            try {
                var fileName = Path.GetFileNameWithoutExtension(pdfFromDoc);
                int pageCount = PdfExtractor.GetPageCount(pdfFromDoc);
                for (int i = 1; i <= pageCount; i++) {
                    var outputPath = Path.Combine(dumpPath, String.Format("{0}_{1}.pdf", fileName, i));
                    if (ExtractPages(pdfFromDoc, outputPath, i, i))
                        output.Add(outputPath);
                }
            }
            catch {
                throw;
            }
            return output;
        }

        public static string PDFToPNG(string pdf) {
            var path = pdf.Replace(".pdf", ".png");
            try {
                var settings = new GhostscriptSharp.GhostscriptSettings();
                settings.Device = GhostscriptSharp.Settings.GhostscriptDevices.png256;
                settings.Resolution = new System.Drawing.Size(72, 72);
                settings.Page.Start = 1;
                settings.Page.End = 1;
                settings.Page.AllPages = false;
                var size = new GhostscriptSharp.Settings.GhostscriptPageSize();
                size.Native = GhostscriptSharp.Settings.GhostscriptPageSizes.letter;
                settings.Size = size;
                GhostscriptSharp.GhostscriptWrapper.GenerateOutput(pdf, path, settings);
                if (!System.IO.File.Exists(path))
                    path = string.Empty;
            }
            catch {
                throw;
            }
            return path;
        }
        
        public static void JoinPDFs(string[] files, string outFile) {
            using (var doc = new Document()) {
                using (var writer = new PdfCopy(doc, new FileStream(outFile, FileMode.Create))) {
                    if (writer == null)
                        return;

                    writer.SetMergeFields();
                    doc.Open();

                    foreach (var file in files) {
                        using (var reader = new PdfReader(file)) {
                            reader.ConsolidateNamedDestinations();

                            writer.AddDocument(reader);
                        }
                    }
                }
            }
        }

        public static void AddPages(string[] files, string outFile) {
            using (var doc = new Document()) {
                using (var writer = new PdfCopy(doc, new FileStream(outFile, FileMode.Create))) {
                    if (writer == null)
                        return;

                    writer.SetMergeFields();
                    doc.Open();

                    foreach (var file in files) {
                        using (var reader = new PdfReader(file)) {
                            reader.ConsolidateNamedDestinations();

                            for (int i = 1; i <= reader.NumberOfPages; i++) {                
                                PdfImportedPage page = writer.GetImportedPage(reader, i);
                                writer.AddPage(page);
                            }
                        }
                    }
                }
            }
        }
    }
}
