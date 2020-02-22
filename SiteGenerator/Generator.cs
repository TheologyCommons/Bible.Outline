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
        public  static void Create(string markDownFile, Action<string> fileCreated, Action<string> indexCreated)
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
                    create(p, rootFileName, addIndex);
                }

            }
            else
            {
                create(outlineParagraph,rootFileName, addIndex);
            }

            createIndex(filesAdded, rootFileName,indexCreated);
            fileCreated(rootFileName);
        }

        private static void createIndex(List<string> filesAdded,string rootFileName, Action<string> indexCreated)
        {
            var currentDirectory = System.IO.Directory.GetCurrentDirectory();
            var targetName = System.IO.Path.GetFileNameWithoutExtension(rootFileName);
            var indexFileTargetDirectory = System.IO.Path.Combine(currentDirectory, "HTML", targetName);
            var indexFileTargetPath = System.IO.Path.Combine(indexFileTargetDirectory, "index.html");
            var writer = System.IO.File.CreateText(indexFileTargetPath);
            foreach (var fileAdded in filesAdded)
            {
                System.Uri indexFileTargetDirectoryUri = new Uri(indexFileTargetDirectory+"/");

                System.Uri fileUri = new Uri(fileAdded);



                Uri relativeUri = indexFileTargetDirectoryUri.MakeRelativeUri(fileUri);
                writer.WriteLine($"<a href='{relativeUri.ToString()}'>{System.IO.Path.GetFileName(fileAdded)}</a><br/>");
            }

            writer.Close();
            indexCreated(indexFileTargetPath);
        }

        private static void create(OutlineParagraph outlineParagraph, string targetFile, Action<string> fileCreated)
        {
            var pptApplication = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();

            var currentDirectory = System.IO.Directory.GetCurrentDirectory();
            var targetName = System.IO.Path.GetFileNameWithoutExtension(targetFile);
            var pptFileTargetPath = System.IO.Path.Combine(currentDirectory, "PPT", targetName, outlineParagraph.Text.Trim() + ".ppt");


            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoFalse);

            var filePath = typeof(Generator).Assembly.Location;
            var path = System.IO.Path.GetDirectoryName(filePath);


            pptPresentation.ApplyTheme(path + @"\BibleStudy.thmx");
            createPresentation(outlineParagraph, pptPresentation,fileCreated);

            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(pptFileTargetPath));

                pptPresentation.SaveAs(pptFileTargetPath);
                fileCreated(pptFileTargetPath);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }

            var odpOutputfile = Path.Combine(currentDirectory, "ODP", targetName, outlineParagraph.Text.Trim()+ ".odp");



            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(odpOutputfile));
            pptPresentation.SaveAs(odpOutputfile, PpSaveAsFileType.ppSaveAsOpenDocumentPresentation);
            fileCreated(odpOutputfile);
            pptPresentation.Close();
            pptApplication.Quit();


            var padHtmlDocfile = Path.Combine(currentDirectory, "HTML", targetName, outlineParagraph.Text.Trim() + ".html");
            createPandoc(outlineParagraph, padHtmlDocfile, fileCreated, "html");


        }

        private static void createPandoc(OutlineParagraph outlineParagraph,string targetFilePath, Action<string> fileCreated, string format)
        {
            string processName = @"C:\Program Files\Pandoc\pandoc.exe";
            string args = String.Format($"-r markdown  -t {format}");

            ProcessStartInfo psi = new ProcessStartInfo(processName, args);

            psi.RedirectStandardOutput = true;
            psi.RedirectStandardInput = true;

            Process p = new Process();
            p.StartInfo = psi;
            psi.UseShellExecute = false;
            p.Start();

            string outputString = "";
            string paraMarkDown = outlineParagraph.MarkDownAsync().Result;
            byte[] inputBuffer = Encoding.UTF8.GetBytes(paraMarkDown);
            p.StandardInput.BaseStream.Write(inputBuffer, 0, inputBuffer.Length);
            p.StandardInput.Close();

            p.WaitForExit(2000);
            using (System.IO.StreamReader sr = new System.IO.StreamReader(
                p.StandardOutput.BaseStream))
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(targetFilePath));
                var fileStream = File.Create(targetFilePath);
                p.StandardOutput.BaseStream.CopyTo(fileStream);
                fileStream.Close();
            }

            fileCreated(targetFilePath);

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