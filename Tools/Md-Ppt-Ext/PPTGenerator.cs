using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Md_Ppt_Ext
{
    class PPTGenerator
    {
        internal static void Create(string markDownFile, string targetFile)
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

            var outlineParagraph = new OutlineParagraph(null, paragraphs);
            createPowerPoint(outlineParagraph, targetFile);


        }

        private void createPowerPoint(OutlineParagraph outlineParagraph, string pptOutput)
        {

        }
    }
}
