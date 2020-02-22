using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiteGenerator
{
    internal class OutlineParagraph
    {
        public OutlineParagraph(Paragraph paragraph,List<Paragraph> children)
        {
            Text = paragraph?.First();
            Level = paragraph?.Level()??0;
            Children = new List<OutlineParagraph>();
            List<Paragraph> grandchildren = new List<Paragraph>();

            Paragraph workingPara = null;


            foreach (var para in children)
            {

                if (para.Level() > 0)
                {
                    if (workingPara == null)
                    {
                        workingPara = para;
                    }
                    else if (para.Level() <= workingPara.Level())
                    {
                        Children.Add(new OutlineParagraph(workingPara,grandchildren));
                        grandchildren = new List<Paragraph>();
                        workingPara = para;
                    }
                    else
                    {
                        grandchildren.Add(para);
                    }
                }
            }

            if (workingPara != null)
            {
                Children.Add(new OutlineParagraph(workingPara, grandchildren));
            }
        }

        public string Text { get; }

        public int Level { get; }

        public List<OutlineParagraph> Children { get; }

        public async Task<string> MarkDownAsync()
        {
            var text = Text;
            bool isHeader = Children.Any() || Level ==1 || Level ==2;
            while (text.StartsWith("#"))
            {
                isHeader = true;
                text = text.Trim('#');
            }

            var builder = new StringBuilder();
            if (isHeader)
            {
                builder.AppendLine($"# {text}");
                await childrenMarkDown(builder,this);
                return builder.ToString();
            }

            return text;
        }

        private async Task childrenMarkDown(StringBuilder builder, OutlineParagraph paragraph,string preface ="#")
        {
            foreach (var child in paragraph.Children)
            {
                using (StringReader reader = new StringReader(await child.MarkDownAsync()))
                {
                    string line = await reader.ReadLineAsync();
                    builder.AppendLine($"{preface}{line}");
                    await childrenMarkDown(builder, child,preface+"#");
                }
            }
        }
    }
}