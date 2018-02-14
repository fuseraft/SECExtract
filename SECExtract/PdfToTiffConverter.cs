using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace SECExtract
{
    public static class PdfToTiffConverter
    {        
        private static ImageCodecInfo GetImageCodecInfoByMimeType(string mimeType)
        {
            return ImageCodecInfo.GetImageEncoders().Where(i => i.MimeType.Equals(mimeType)).SingleOrDefault();
        }

        /// <summary>
        /// Builds a multi-frame tiff from a multi-page pdf
        /// </summary>
        /// <param name="pdf">Path to multi-page pdf file</param>
        /// <returns>Path to a single tiff image file containing multiple pages from a pdf</returns>
        public static string BuildMultiFrameTiff(string pdf, string outputPath)
        {
            // get each page as individual pdfs and bitmaps for each page
            var pdfPages = GetPages(pdf, outputPath);
            var pageImages = GetPageImages(pdfPages);

            // this is where we'll write our final multi-frame tiff
            var tiffOutputPath = Path.Combine(outputPath, $"{Path.GetFileNameWithoutExtension(pdf)}.tiff");

            // if there are bitmaps to process
            if (pageImages.Count > 0)
            {
                Image tiff = null;

                // start with the first png
                var bmp = (Bitmap)Image.FromFile(pageImages[0]);
                var encoder = GetImageCodecInfoByMimeType("image/tiff");

                // keep list of memory streams to close later
                // we have to keep the memory streams open until we're done with image processing
                var msPages = new List<MemoryStream>();
                var msFirstPage = new MemoryStream();

                // save the first page as a tiff into our first memory stream
                bmp.Save(msFirstPage, ImageFormat.Tiff);
                // create an object from our tiff stream
                tiff = Image.FromStream(msFirstPage);

                // add to list of mem streams to close later
                msPages.Add(msFirstPage);

                // create an encoder parameter for mutli frame tiff
                var encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(
                    System.Drawing.Imaging.Encoder.SaveFlag,
                    (long)EncoderValue.MultiFrame
                );

                // save tiff to output path using our encoder parameter and mime type
                tiff.Save(tiffOutputPath, encoder, encoderParams);

                // if there are multiple pages, do the same for each page
                if (pageImages.Count > 1)
                {
                    for (var i = 1; i < pageImages.Count; i++)
                    {
                        // this frame will be added to the page dimension of the tiff
                        var frameEncoderParams = new EncoderParameters(1);
                        frameEncoderParams.Param[0] = new EncoderParameter(
                            System.Drawing.Imaging.Encoder.SaveFlag,
                            (long)EncoderValue.FrameDimensionPage
                        );

                        // create new mem stream for current page
                        var msPage = new MemoryStream();

                        // get bitmap from individual page pdf, save to mem stream, and save to tiff object using frame parameter
                        var tiffPage = (Bitmap)Image.FromFile(pageImages[i]);
                        tiffPage.Save(msPage, ImageFormat.Tiff);
                        tiff.SaveAdd(tiffPage, frameEncoderParams);

                        // add to list of mem streams for closing later
                        msPages.Add(msPage);
                    }
                }

                // flush the bytes to the tiff
                var saveEncoderParams = new EncoderParameters(1);
                saveEncoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);

                tiff.SaveAdd(saveEncoderParams);

                // close each memory stream
                foreach (var ms in msPages)
                {
                    ms.Close();
                }
            }

            // remove temporary pdf pages and bitmaps
            Cleanup(pdfPages, pageImages);

            // return path to final tiff
            return tiffOutputPath;
        }

        private static void Cleanup(params List<string>[] fileLists)
        {
            foreach (var fileList in fileLists)
            {
                foreach (var file in fileList)
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch { }
                }
            }
        }

        private static List<string> GetPageImages(List<string> pdfPages)
        {
            var pages = new List<string>();

            foreach (var page in pdfPages)
            {
                var png = SECExtract.PdfExtractor.PdfToImage(
                    page,
                    SECExtract.PdfExtractor.GSImageFormat.Bmp,
                    new System.Drawing.Size
                    {
                        Width = 150,
                        Height = 150
                    }
                );

                pages.Add(png);
            }

            return pages;
        }

        private static List<string> GetPages(string pdf, string outputPath)
        {
            var pages = new List<string>();
            var count = 0;

            foreach (var page in PdfExtractor.SplitPdf(pdf, outputPath))
            {
                var tmp = $"Temp{DateTime.Now.ToString("yyyyMMddhhmmsstt")}_{count}.pdf";
                File.Copy(page, tmp);
                pages.Add(tmp);
                File.Delete(page);
                ++count;
            }

            return pages;
        }
    }
}
