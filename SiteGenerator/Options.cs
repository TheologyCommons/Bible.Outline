using System;
using System.Collections.Generic;
using System.Text;
using CommandLine;

namespace SiteGenerator
{
    public class Options
    {
        [Option('m', "MarkDown", Required = true, HelpText = "Mark Down Directory")]
        public string MarkDownDirectory { get; set; }

        [Option('o', "Output", Required = true, HelpText = "Folder to create the site in")]
        public string Output { get; set; }
    }
}
