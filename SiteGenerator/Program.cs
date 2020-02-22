using System;
using CommandLine;

namespace SiteGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var processor = new FileProcessor(new Logger());
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o => { processor.Create(o.MarkDownDirectory,o.Output); }
                );
        }
    }
}
