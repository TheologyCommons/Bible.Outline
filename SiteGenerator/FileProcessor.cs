using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SiteGenerator
{
    class FileProcessor
    {
        private Logger _logger;
        private Generator _pptGenerator;

        public FileProcessor(Logger logger)
        {
            _logger = logger;
            _pptGenerator = new Generator();
        }
        internal void Create(string markDownDirectory,string outputFolder = ".")
        {
            if (!System.IO.Directory.Exists(markDownDirectory))
            {
                throw new ArgumentException("Markdown directory does not exist");
            }
            if (!System.IO.Directory.Exists(outputFolder))
            {
                throw new ArgumentException("Output directory does not exist");
            }

            System.IO.Directory.SetCurrentDirectory(outputFolder);
            processFiles(markDownDirectory);
            createMenu(outputFolder);

        }

        private void createMenu(string outputFolder)
        {
            var markDownDirectory = System.IO.Path.Combine(outputFolder, "MD");
            System.IO.Directory.CreateDirectory(markDownDirectory);

            var currentDirectory = System.IO.Directory.GetCurrentDirectory();
            var indexFileTargetPath = System.IO.Path.Combine(currentDirectory,"../","_includes", "menu.html");
            var writer = System.IO.File.CreateText(indexFileTargetPath);
            writer.WriteLine("<!-- Menu.html -->");
            var q = new Queue<string>();
            writer.WriteLine("<nav id=\"ml-menu\" class=\"menu\">");
            writer.WriteLine("<button class=\"action action--close\" aria-label=\"Close Menu\"><span class=\"icon icon--cross\"></span></button>");
            writer.WriteLine("<div class=\"menu__wrap\">");
            writer.Write(buildDivForDirectory(markDownDirectory,"main",(directoryPath)=>q.Enqueue(directoryPath)));
            while (q.Any())
            {
                string directory = q.Dequeue();
                writer.Write(buildDivForDirectory(directory, nameFromPath(directory), (directoryPath) => q.Enqueue(directoryPath)));
            }
            writer.WriteLine("</div>");
            writer.WriteLine("</nav>");
            writer.Close();
        }

        private static string buildDivForDirectory(string markDownDirectory, string dataMenu,Action<string> directoryLocated)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"<ul data-menu=\"{dataMenu}\" class=\"menu__level\" tabindex=\" - 1\" role=\"menu\" aria-label=\"{formatMenuName(System.IO.Path.GetFileNameWithoutExtension(markDownDirectory))}\">");
            foreach (var file in System.IO.Directory.GetFiles(markDownDirectory).Where(filename => !filename.EndsWith(".git")))
            {
                builder.AppendLine(
                    $"<li class=\"menu__item\" role=\"menuitem\"><a class=\"menu__link\" href=\"/{pathForFile(file)}\">{System.IO.Path.GetFileNameWithoutExtension(file)}</a></li>");
            }

            foreach (var directory in System.IO.Directory.GetDirectories(markDownDirectory))
            {
                builder.AppendLine(
                    $"<li class=\"menu__item\" role=\"menuitem\"><a class=\"menu__link\" data-submenu=\"{nameFromPath(directory)}\" aria-owns=\"{nameFromPath(directory)}\" href=\"#\">{formatMenuName(System.IO.Path.GetFileNameWithoutExtension(directory))}</a></li>");
                directoryLocated(directory);
            }
            builder.AppendLine("</ul>");
            return builder.ToString();
        }

        private static string pathForFile(string file)
        {
            if (file.EndsWith(".markdown"))
            {
                file = file.Replace(".markdown", ".html");
            }
            Uri currentDirectory = new Uri(System.IO.Directory.GetCurrentDirectory());
            Uri targetFile = new Uri(file);
            var relative = currentDirectory.MakeRelativeUri(targetFile);
            
            return relative.ToString();
        }

        private static string formatMenuName(string name)
        {
            if (name.Length > 15)
            {
                return name.Substring(0, 10) + "...";
            }
            else
            {
                return name;
            }
        }

        private static string nameFromPath(string directory)
        {
            return directory
                .Replace("\\", "_")
                .Replace(" ", "_")
                .Replace(":","");
        }

        private void processFiles(string markDownDirectory)
        {
            List<string> indexesAdded = new List<string>();
            foreach (var file in System.IO.Directory.EnumerateFiles(markDownDirectory))
            {
                processFile(markDownDirectory, file, (fileName) =>
                {
                    indexesAdded.Add(fileName);
                });
            }

            foreach (var directory in System.IO.Directory.EnumerateDirectories(markDownDirectory))
            {
                processFiles(directory);
            }
        }

        private void processFile(string directory, string file,Action<string> indexCreated)
        {
            if (file.EndsWith(".md")&&System.IO.File.Exists(file))
            {
                var fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                Generator.Create(file,
                    (generatedFileName)=>fileCreated(generatedFileName), indexCreated
                    );
            }
        }

        private void fileCreated(string fileName)
        {
            _logger.Log(fileName);
        }
    }
}
