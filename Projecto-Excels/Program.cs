using NDesk.Options;
using System.Collections.Generic;
using System;

namespace Projecto_Excels
{
    class Program
    {

        private const string SHEET_ID = "16BRJ-GuURF_rIKPosekyEXehGGQxIgGK5PLiCKOBQSQ";

        static void Main(string[] args)
        {
            string CURRENT_DIR = System.IO.Directory.GetParent(@"./").Parent.Parent.Parent.Parent.FullName;
            string AUTH_FILE = $@"{CURRENT_DIR}\Projecto-Excels\auth.json";
            string OUTPUT_FILE = $@"{CURRENT_DIR}\Projecto-Excels\results.xls";

            var options = Program.GetOptions(args);
            var gsh = new GoogleDiscourse(AUTH_FILE, SHEET_ID);
            var gsp = new GoogleSheetParameters() {
                FromCell = options["fromCell"],
                ToCell = options["toCell"],
                SheetName = options["sheetName"]
            };
            var rowValues = gsh.GetDataFromSheet(gsp);
            var ed = new ExcelDiscourse(OUTPUT_FILE);
            ed.WriteData(rowValues);
            ed.SaveToFile();
        }

        static Dictionary<string, string> GetOptions(string[] args)
        {
            Dictionary<string, string> options = new Dictionary<string, string>();
            var optionSet = new OptionSet()
            {
                { "s|sheet=", "Sheet name", value => options.Add("sheetName", value) },
                { "f|from=",  "The start sheet cell", value => options.Add("fromCell", value) },
                { "t|to=", "The end sheet cell", value => options.Add("toCell", value) },
            };

            try
            {
                optionSet.Parse(args);
            } catch (OptionException e)
            {
                Program.ShowHelp(optionSet);
                throw new InvalidOperationException($"Invalid arguments {e.Message}");

            }
            if(!options.ContainsKey("fromCell"))
            {
                Program.ShowHelp(optionSet);
                throw new InvalidOperationException("Missing required option f|from");
            }
            else if(!options.ContainsKey("toCell"))
            {
                Program.ShowHelp(optionSet);
                throw new InvalidOperationException("Missing required option t|to");
            }
            else if(!options.ContainsKey("sheetName"))
            {
                Program.ShowHelp(optionSet);
                throw new InvalidOperationException("Missing required option s|sheet");
            }
            return options;
        }

        static void ShowHelp(OptionSet optionSet)
        {
            Console.WriteLine("Usage: <binary> -s [SHEET_NAME] -f [FROM_CELL] -t [TO_CELL]");
            optionSet.WriteOptionDescriptions(Console.Out);
        }
    } 
}

