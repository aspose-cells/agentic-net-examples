using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadTabDelimitedFile
    {
        public static void Run()
        {
            string tsvPath = "input.tsv";

            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
            loadOptions.Separator = '\t';
            loadOptions.ConvertNumericData = true;
            loadOptions.ConvertDateTimeData = true;
            loadOptions.HasTextQualifier = true;

            Workbook workbook = new Workbook(tsvPath, loadOptions);

            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            Console.WriteLine($"A1: Value = {cells["A1"].Value}, Type = {cells["A1"].Type}");
            Console.WriteLine($"B1: Value = {cells["B1"].Value}, Type = {cells["B1"].Type}");
            Console.WriteLine($"C1: Value = {cells["C1"].Value}, Type = {cells["C1"].Type}");

            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadTabDelimitedFile.Run();
        }
    }
}