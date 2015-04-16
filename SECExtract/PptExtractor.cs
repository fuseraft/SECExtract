using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using PKG = DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Text;

namespace SECExtract {
    public class PptExtractor {
        const string PNG_FORMAT = @".\Dump\{1}_{2}.png";
        const string TXT_FORMAT = @".\Dump\{1}_{2}.txt";

        public static int SlideCount(string pathToPPT) {
            int slideCount = 0;
            try {
                using (var docPPT = PKG.PresentationDocument.Open(pathToPPT, false)) {
                    return SlideCount(docPPT);
                }
            }
            catch {
                slideCount--;
            }
            return slideCount;
        }

        public static int SlideCount(PKG.PresentationDocument docPPT) {
            if (docPPT == null)
                throw new ArgumentNullException("docPPT");

            int count = 0;

            var part = docPPT.PresentationPart;
            if (part != null)
                count = part.SlideParts.Count();

            return count;
        }

        public static string GetSlideText(string pathToPPT, int index) {
            var slideText = string.Empty;
            try {
                using (var ppt = PKG.PresentationDocument.Open(pathToPPT, false)) {
                    var part = ppt.PresentationPart;
                    PKG.SlidePart slide = (PKG.SlidePart)part.GetPartById((part.Presentation.SlideIdList.ChildElements[index] as SlideId).RelationshipId);

                    var sb = new StringBuilder();
                    foreach (var text in slide.Slide.Descendants<A.Text>())
                        sb.Append(text.Text);
                    slideText = sb.ToString();
                }
            }
            catch {
                throw;
            }
            return slideText;
        }

        public static string ReadCommentFromPresentation(string file, int index) {
            var slideComment = string.Empty;
            try {
                using (PKG.PresentationDocument doc = PKG.PresentationDocument.Open(file, true)) {
                    PKG.SlidePart part = GetSlideByIndex(doc, index);
                    PKG.SlideCommentsPart comments;

                    if (part.GetPartsOfType<PKG.SlideCommentsPart>().Count() == 0)
                        comments = part.AddNewPart<PKG.SlideCommentsPart>();
                    else
                        comments = part.GetPartsOfType<PKG.SlideCommentsPart>().First();

                    if (comments.CommentList == null)
                        comments.CommentList = new CommentList();

                    var sb = new StringBuilder();
                    foreach (var comment in comments.CommentList)
                        sb.Append(comment.InnerText); // InnerXML
                    slideComment = sb.ToString();
                }
            }
            catch {
                throw;
            }
            return slideComment;
        }

        public static void RemoveCommentsFromPresentation(string file, int index) {
            try {
                using (PKG.PresentationDocument doc = PKG.PresentationDocument.Open(file, true)) {
                    PKG.SlidePart part = GetSlideByIndex(doc, index);

                    if (part.GetPartsOfType<PKG.SlideCommentsPart>().Count() > 0) {
                        var comments = part.GetPartsOfType<PKG.SlideCommentsPart>();
                        part.DeleteParts<PKG.SlideCommentsPart>(comments);
                    }
                }
            } catch {
                throw;
            }
        }

        public static void AddCommentToPresentation(string file, string initials, string name, string text, int index) {
            try {
                using (PKG.PresentationDocument doc = PKG.PresentationDocument.Open(file, true)) {
                    PKG.CommentAuthorsPart authorsPart;

                    if (doc.PresentationPart.CommentAuthorsPart == null)
                        authorsPart = doc.PresentationPart.AddNewPart<PKG.CommentAuthorsPart>();
                    else
                        authorsPart = doc.PresentationPart.CommentAuthorsPart;

                    if (authorsPart.CommentAuthorList == null)
                        authorsPart.CommentAuthorList = new CommentAuthorList();

                    uint authorId = 0;
                    CommentAuthor author = null;

                    if (authorsPart.CommentAuthorList.HasChildren) {
                        var authors = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Where(a => a.Name == name && a.Initials == initials);

                        if (authors.Any()) {
                            author = authors.First();
                            authorId = author.Id;
                        }

                        if (author == null)
                            authorId = authorsPart.CommentAuthorList.Elements<CommentAuthor>().Select(a => a.Id.Value).Max();
                    }

                    if (author == null) {
                        authorId++;

                        author = authorsPart.CommentAuthorList.AppendChild<CommentAuthor>(
                            new CommentAuthor() {
                                Id = authorId,
                                Name = name,
                                Initials = initials,
                                ColorIndex = 0
                            }
                        );
                    }

                    PKG.SlidePart slidePart1 = GetSlideByIndex(doc, index);

                    PKG.SlideCommentsPart commentsPart;

                    if (slidePart1.GetPartsOfType<PKG.SlideCommentsPart>().Count() == 0)
                        commentsPart = slidePart1.AddNewPart<PKG.SlideCommentsPart>();
                    else
                        commentsPart = slidePart1.GetPartsOfType<PKG.SlideCommentsPart>().First();

                    if (commentsPart.CommentList == null)
                        commentsPart.CommentList = new CommentList();

                    uint commentIdx = author.LastIndex == null ? 1 : author.LastIndex + 1;
                    author.LastIndex = commentIdx;

                    DocumentFormat.OpenXml.Spreadsheet.Comment comment = commentsPart.CommentList.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Comment>(
                        new DocumentFormat.OpenXml.Spreadsheet.Comment() {
                            AuthorId = authorId
                        }
                    );

                    comment.Append(
                        new Position() { X = 100, Y = 200 },
                        new Text() { Text = text }
                    );

                    authorsPart.CommentAuthorList.Save();

                    commentsPart.CommentList.Save();
                }
            }
            catch {
                // PowerPoint is probably running
                throw;
            }
        }

        public static PKG.SlidePart GetSlideByIndex(PKG.PresentationDocument ppt, int index) {
            var part = ppt.PresentationPart;
            PKG.SlidePart slide = (PKG.SlidePart)part.GetPartById((part.Presentation.SlideIdList.ChildElements[index] as SlideId).RelationshipId);

            return slide;
        }

        public static List<string> PPT2TXT(string pathToPPT, string outputPath) {
            var input = System.IO.Path.GetFileNameWithoutExtension(pathToPPT);
            var paths = new List<string>();

            int slideCount = SlideCount(pathToPPT);
            for (int i = 0; i < slideCount; i++) {
                var output = String.Format(TXT_FORMAT, outputPath, input, i + 1);
                var slideText = ReadCommentFromPresentation(pathToPPT, i);

                File.WriteAllText(output, slideText.Replace("'", "''"));
                if (File.Exists(output))
                    paths.Add(output);
            }

            return paths;
        }

        public static List<string> PPT2PNG(string pathToPPT, string outputPath) {
            return PPT2PNG(pathToPPT, outputPath, 1024, 768);
        }

        public static List<string> PPT2PNG(string pathToPPT, string outputPath, int width, int height) {
            var input = System.IO.Path.GetFileNameWithoutExtension(pathToPPT);
            var paths = new List<string>();
            Microsoft.Office.Interop.PowerPoint._Application app = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = app.Presentations.Open2007(pathToPPT, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            int i = 1;
            foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in presentation.Slides) {
                var output = String.Format(PNG_FORMAT, outputPath, input, i);
                slide.Export(output, "PNG", width, height);
                i++;
                if (File.Exists(output))
                    paths.Add(output);
            }
            return paths;
        }
    }
}