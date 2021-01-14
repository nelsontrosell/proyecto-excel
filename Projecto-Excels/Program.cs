using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;

namespace Projecto_Excels
{
    class Program
    {
        private const string AUTH_FILE = "c:\\Users\\User\\Downloads\\Projecto-Excels\\Projecto-Excels\\auth.json";
        private const string SHEET_ID = "16BRJ-GuURF_rIKPosekyEXehGGQxIgGK5PLiCKOBQSQ";
        private const string OUTPUT_FILE = "c:\\Users\\User\\Downloads\\Projecto-Excels\\Projecto-Excels\\results.xls";

        static void Main(string[] args)
        {
            var gsh = new GoogleDiscourse(AUTH_FILE, SHEET_ID);
            var gsp = new GoogleSheetParameters() {
                RangeColumnStart = "A",
                RangeRowStart = 1,
                RangeColumnEnd = "N",
                RangeRowEnd = 13,
                FirstRowIsHeaders = false,
                SheetName = "DAISY MENDEZ"
            };
            var rowValues = gsh.GetDataFromSheet(gsp);
            var ed = new ExcelDiscourse(OUTPUT_FILE);
            ed.WriteData(rowValues);
            ed.SaveToFile();
        }
    } 
}

