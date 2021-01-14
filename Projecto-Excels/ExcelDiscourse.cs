using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;

namespace Projecto_Excels
{
    class ExcelDiscourse
    {

        private readonly Workbook _workbook;
        private readonly string _path;
        private readonly FileStream _stream;

        public ExcelDiscourse(string path)
        {
            _path = path;
            _workbook = new Workbook();
            _stream = new FileStream(_path, FileMode.Create);
        }

        public void WriteData(IList<IList<Object>> cells)
        {
            Worksheet sheet = _workbook.Worksheets[0];

            for(int i = 0; i < cells.Count; i++)
            {
                IList<Object> row = cells[i];
                for(int j = 0; j < row.Count; j++)
                {
                      var cellName = GetColumnName(j) + (i+1);
                    sheet.Range[cellName].Text = (string)row[j];
                }
            }
        }

        public void SaveToFile()
        {
            _workbook.SaveToStream(_stream);
        } 

        private string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];
            return value;
        }
    }
}
 