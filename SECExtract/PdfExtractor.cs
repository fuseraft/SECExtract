using System;
using System.Collections.Generic;
using System.Text;

using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace SECExtract {
    public class PdfExtractor {
        public static int GetPageCount(string sourcePdfPath) {
            var pageCount = 0;

            using (var reader = new PdfReader(sourcePdfPath)) {
                pageCount = reader.NumberOfPages;
            }
            
            return pageCount;
        }

        public static string ReadPdfFile(string fileName, int startPage = 1, int endPage = -1) {
            var text = new StringBuilder();

            if (File.Exists(fileName)) {
                using (var pdfReader = new PdfReader(fileName)) {
                    if (startPage < 1 || startPage > endPage)
                        startPage = 1;

                    if (endPage < 0 || endPage < startPage)
                        endPage = pdfReader.NumberOfPages;

                    for (int page = startPage; page <= endPage; page++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                        currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                        text.Append(currentText);
                    }
                }
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
                pdfCopyProvider = new PdfCopy(sourceDocument, new FileStream(outputPdfPath, FileMode.Create));

                sourceDocument.Open();

                for (int i = startPage; i <= endPage; i++) {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }

                sourceDocument.Close();
                reader.Close();
                success = true;
            }
            catch { throw; }

            return success;
        }

        public static List<string> SplitPdf(string pdfFromDoc, string dumpPath) {
            var output = new List<string>();

            try {
                var fileName = Path.GetFileNameWithoutExtension(pdfFromDoc);
                int pageCount = GetPageCount(pdfFromDoc);
                for (int i = 1; i <= pageCount; i++) {
                    var outputPath = Path.Combine(dumpPath, $"{fileName}_{i}.pdf");
                    if (ExtractPages(pdfFromDoc, outputPath, i, i))
                        output.Add(outputPath);
                }
            }
            catch { throw; }

            return output;
        }

        public enum GSImageFormat {
            Bmp = GhostscriptSharp.Settings.GhostscriptDevices.bmp256,
            Jpeg = GhostscriptSharp.Settings.GhostscriptDevices.jpeg,
            Png = GhostscriptSharp.Settings.GhostscriptDevices.png256,
            Tiff = GhostscriptSharp.Settings.GhostscriptDevices.tiff32nc
        }

        private static string GetExtensionFromImageFormat(GSImageFormat imageFormat) {
            var ext = string.Empty;

            switch (imageFormat) {
                case GSImageFormat.Bmp:
                    ext = ".bmp";
                    break;
                case GSImageFormat.Jpeg:
                    ext = ".jpg";
                    break;
                case GSImageFormat.Png:
                    ext = ".png";
                    break;
                case GSImageFormat.Tiff:
                    ext = ".tiff";
                    break;
            }

            return ext;
        }

        public static string PdfToImage(string pdf, GSImageFormat imageFormat, System.Drawing.Size overrideSize) {
            var ext = GetExtensionFromImageFormat(imageFormat);

            if (string.IsNullOrEmpty(ext))
                throw new Exception($"Invalid ImageFormat: {imageFormat}");

            var path = pdf.Replace(".pdf", ext);

            try {
                var settings = new GhostscriptSharp.GhostscriptSettings()
                {
                    Device = (GhostscriptSharp.Settings.GhostscriptDevices)imageFormat,
                    Resolution = overrideSize,
                    Page = {
                        Start = 1,
                        End = 1,
                        AllPages = false
                    },
                    Size = new GhostscriptSharp.Settings.GhostscriptPageSize()
                    {
                        Native = GhostscriptSharp.Settings.GhostscriptPageSizes.letter
                    }
                };

                GhostscriptSharp.GhostscriptWrapper.GenerateOutput(pdf, path, settings);

                if (!File.Exists(path))
                    throw new Exception($"Could not generate output: {path}");
            }
            catch { throw; }

            return path;
        }

        public static byte[] GetFlatPdf(string pdf)
        {
            PdfReader.unethicalreading = true;
            using (var reader = new PdfReader(pdf))
            using (var ms = new MemoryStream())
            using (var stamper = new PdfStamper(reader, ms) { FormFlattening = true })
            {
                PdfReader.unethicalreading = true;
                return ms.ToArray();
            }
        }

        public static byte[] GetFlattenedPdfBytes(string sourcePdfPath) {
            var bytes = new List<byte>();
            
            if (File.Exists(sourcePdfPath)) {
                using (var ms = new MemoryStream())
                using (var reader = new PdfReader(sourcePdfPath))
                {
                    //PdfReader.unethicalreading = true;
                    using (var stamper = new PdfStamper(reader, ms) { FormFlattening = true })
                    {
                        bytes.AddRange(ms.ToArray());
                    }
                }
            }

            return bytes.ToArray();
        }
        
        public static void JoinPdfs(string[] files, string outFile) {
            using (var ms = new MemoryStream()) {
                var doc = new Document();
                var copy = new PdfSmartCopy(doc, ms);
                var readers = new List<PdfReader>();

                doc.Open();

                foreach (var file in files) {
                    var reader = new PdfReader(file);
                    reader.ConsolidateNamedDestinations();
                    copy.AddDocument(reader);
                    readers.Add(reader);
                }

                copy.Close();

                foreach (var reader in readers)
                    reader.Close();

                doc.Close();

                ms.Flush();
                File.WriteAllBytes(outFile, ms.ToArray());
            }
        }

        public static void AddPages(string[] files, string outFile) {
            using (var doc = new Document())
            using (var writer = new PdfCopy(doc, new FileStream(outFile, FileMode.Create))) {
                if (writer == null) return;

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
