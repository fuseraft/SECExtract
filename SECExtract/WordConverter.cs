using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Microsoft.Office.Interop.Word;

namespace SECExtract {
    public class WordConverter {
        public static string DocToPDF(string inputPath, string outputPath) {
            var returnValue = string.Empty;
            var path = Path.GetFileName(inputPath);
            var extension = ".pdf";
            try {
                var word = new Microsoft.Office.Interop.Word.Application();
                word.Visible = false;

                object input = (object)inputPath;
                object output = (object)Path.Combine(outputPath, path.Replace(".doc", extension));
                object outFormat = (object)WdSaveFormat.wdFormatPDF;
                var doc = word.Documents.Open(ref input);
                doc.SaveAs2(ref output, ref outFormat);
                ((Document)doc).Close();
                word.Quit();
                returnValue = (string)output;
            }
            catch (Exception) {
            }

            return returnValue;
        }
    }
}