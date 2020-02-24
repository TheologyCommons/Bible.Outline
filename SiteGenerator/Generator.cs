using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace SiteGenerator
{
    public class Generator
    {
        public  static void Create(string markDownFile, Action<string> fileCreated, Action<string> indexCreated,Logger logger)
        {
            var paragraphs = new List<Paragraph>();
            using (StreamReader fs = new StreamReader(markDownFile))
            {
                var paragraph = new Paragraph();
                while (!fs.EndOfStream)
                {
                    var line = fs.ReadLine();
                    if (line.All(c => c == 13))
                    {
                        paragraphs.Add(paragraph);
                        paragraph = new Paragraph();
                    }
                    else if (!string.IsNullOrWhiteSpace(line))
                    {
                        paragraph.Add(line);
                    }
                }
            }

            var rootFileName = System.IO.Path.GetFileNameWithoutExtension(markDownFile);
            var para = new Paragraph(){ rootFileName };
            var outlineParagraph = new OutlineParagraph(para, paragraphs);

            List<string> filesAdded = new List<string>();
            Action<string> addIndex = (filename) =>
            {
                filesAdded.Add(filename);
            };
            if (outlineParagraph.Children.Count > 10)
            {
                foreach (var p in outlineParagraph.Children)
                {
                    createPresentation(p, rootFileName, addIndex,logger);
                    createMDFile(outlineParagraph, rootFileName);
                }

            }
            else
            {
                createPresentation(outlineParagraph,rootFileName, addIndex,logger);
                createMDFile(outlineParagraph, rootFileName);
            }
            fileCreated(rootFileName);
        }


        private static void createPresentation(OutlineParagraph outlineParagraph, string targetFile, Action<string> fileCreated, Logger logger)
        {
            var pptApplication = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();

            var currentDirectory = System.IO.Directory.GetCurrentDirectory();
            var targetName = System.IO.Path.GetFileNameWithoutExtension(targetFile);
            //var pptFileTargetPath = System.IO.Path.Combine(currentDirectory, "PPT", targetName, outlineParagraph.Text.Trim() + ".ppt");
            var odpOutputfile = Path.Combine(currentDirectory, "ODP", targetName, outlineParagraph.Text.Trim() + ".odp");
            if (System.IO.File.Exists(odpOutputfile))
            {
                logger.Log($"{odpOutputfile} already exists to recreate it delete the existing version.");
                return;
            }

            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoFalse);

            var filePath = typeof(Generator).Assembly.Location;
            var path = System.IO.Path.GetDirectoryName(filePath);


            pptPresentation.ApplyTheme(path + @"\BibleStudy.thmx");
            createPresentation(outlineParagraph, pptPresentation,fileCreated);

            //try
            //{
            //    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(pptFileTargetPath));

            //    pptPresentation.SaveAs(pptFileTargetPath);
            //    fileCreated(pptFileTargetPath);
            //}
            //catch(Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //    throw;
            //}

            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(odpOutputfile));

            pptPresentation.SaveAs(odpOutputfile, PpSaveAsFileType.ppSaveAsOpenDocumentPresentation);
            fileCreated(odpOutputfile);
            pptPresentation.Close();
            pptApplication.Quit();



        }

        private static void createMDFile(OutlineParagraph outlineParagraph, string targetFile)
        {
            var targetName = System.IO.Path.GetFileNameWithoutExtension(targetFile);

            var Odpfile = Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                "ODP", targetName, outlineParagraph.Text.Trim() + ".odp");
            var MDDocfile = Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                "MD", targetName, outlineParagraph.Text.Trim() + ".markdown");
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(MDDocfile));
            string paraMarkDown = outlineParagraph.MarkDownAsync().Result;
            var writer = System.IO.File.CreateText(MDDocfile);
            writer.WriteLine("---");
            writer.WriteLine("layout: outline");
            writer.WriteLine($"title: {outlineParagraph.Text.Trim()}");
            writer.WriteLine($"presentation: {relativeToCurrent(Odpfile)}");

            writer.WriteLine("---");
            writer.Write(paraMarkDown);
            writer.Close();
        }

        private static object relativeToCurrent(string odpfile)
        {
            var currentUri = new Uri(System.IO.Directory.GetCurrentDirectory());
            var targetUri = new Uri(odpfile);
            return currentUri.MakeRelativeUri(targetUri);
        }

        private static void createPresentation(OutlineParagraph outlineParagraph, Presentation pptPresentation, Action<string> fileCreated)
        {
            var text = outlineParagraph.Text;
            var subTexts = new List<string>();
            if (outlineParagraph.Level == 0)
                createLayoutTitle(text, subTexts, pptPresentation);
            foreach (var child in outlineParagraph.Children)
            {
                if (outlineParagraph.Level > 0) 
                    subTexts.Add(cleanUpText(child.Text));
                if (child.Children.Any())
                {
                    if (outlineParagraph.Level > 0)
                    {
                        createLayoutTitle(text, subTexts, pptPresentation);
                    }

                    createPresentation(child, pptPresentation, fileCreated);
                }
            }
            createLayoutTitle(text, subTexts, pptPresentation);

        }

        private static void createLayoutTitle(string text, List<string> subTexts, Presentation pptPresentation)
        {
            var processedText = cleanUpText(text);
            if (!String.IsNullOrWhiteSpace(text))
            {
                int newSlideNumber = (pptPresentation.Slides.Count + 1);
                var layout = PpSlideLayout.ppLayoutTitle;
                if (subTexts.Count >= 2)
                {
                    layout = PpSlideLayout.ppLayoutText;
                }
                var subTextBuilder = new StringBuilder();
                var subTextBuilder2 = new StringBuilder();

                if (subTexts.Count > 6)
                {
                    layout = PpSlideLayout.ppLayoutTwoColumnText;
                    for (int i = 0; i < subTexts.Count; i++)
                    {
                        if (i < subTexts.Count / 2)
                        {
                            subTextBuilder.AppendLine(subTexts[i]);
                        }
                        else
                        {
                            subTextBuilder2.AppendLine(subTexts[i]);
                        }
                    }
                }
                else
                {
                    foreach (var subText in subTexts)
                    {
                        subTextBuilder.AppendLine(subText);
                    }
                }

                var slide = pptPresentation.Slides.Add(newSlideNumber, layout);
                slide.Shapes[1].TextFrame.TextRange.Text = processedText;
                slide.Shapes[2].TextFrame.TextRange.Text = subTextBuilder.ToString();
                if (subTextBuilder2.Length != 0)
                {
                    slide.Shapes[3].TextFrame.TextRange.Text = subTextBuilder2.ToString();
                }
            }
        }

        private static string cleanUpText(string text)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {

                text = text.Replace(@"\'", "'");
                text = text.Replace("\\\"", "\"");
                text = text.Trim('#');
            }

            return text;
        }
    }
}