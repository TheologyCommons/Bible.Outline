using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

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

        }
        private static void createIndex(List<string> filesAdded)
        {
            var currentDirectory = System.IO.Directory.GetCurrentDirectory();
            var indexFileTargetPath = System.IO.Path.Combine(currentDirectory, "index.html");
            var writer = System.IO.File.CreateText(indexFileTargetPath);
            foreach (var fileAdded in filesAdded)
            {
                System.Uri indexFileTargetDirectoryUri = new Uri(currentDirectory + "/");

                System.Uri fileUri = new Uri(fileAdded);



                Uri relativeUri = indexFileTargetDirectoryUri.MakeRelativeUri(fileUri);

                var directoryName = System.IO.Path.GetDirectoryName(fileAdded);
                var FolderName = new System.IO.DirectoryInfo(directoryName).Name;

                writer.WriteLine($"<a href='{relativeUri.ToString()}'>{FolderName}</a><br/>");
            }

            writer.Close();
            
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

            createIndex(indexesAdded);
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
