using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class WorkbookConverter
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string csvPath = "output.csv";
            string tsvPath = "output.tsv";
            string txtPath = "output.txt";

            try
            {
                Workbook wb = new Workbook(sourcePath);

                wb.Save(csvPath, SaveFormat.CSV);
                Console.WriteLine($"Converted to CSV: {csvPath}");

                wb.Save(tsvPath, SaveFormat.TSV);
                Console.WriteLine($"Converted to TSV: {tsvPath}");

                wb.Save(txtPath, SaveFormat.TabDelimited);
                Console.WriteLine($"Converted to TXT: {txtPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}