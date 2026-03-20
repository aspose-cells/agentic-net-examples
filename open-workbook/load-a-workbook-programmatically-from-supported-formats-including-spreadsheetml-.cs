using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Paths to sample files in different supported formats
            string spreadsheetMlPath = "sample.xml";   // SpreadsheetML (Excel 2003 XML)
            string csvPath = "sample.csv";            // CSV
            string tsvPath = "sample.tsv";            // TSV
            string odsPath = "sample.ods";            // ODS

            // Load SpreadsheetML file
            LoadWorkbook(spreadsheetMlPath, LoadFormat.SpreadsheetML);

            // Load CSV file
            LoadWorkbook(csvPath, LoadFormat.Csv);

            // Load TSV file
            LoadWorkbook(tsvPath, LoadFormat.Tsv);

            // Load ODS file
            LoadWorkbook(odsPath, LoadFormat.Ods);
        }

        private static void LoadWorkbook(string filePath, LoadFormat loadFormat)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            LoadOptions loadOptions = GetLoadOptions(loadFormat);

            try
            {
                using (Workbook workbook = new Workbook(filePath, loadOptions))
                {
                    Console.WriteLine($"Loaded '{Path.GetFileName(filePath)}' successfully.");
                    Console.WriteLine($" - Detected LoadFormat: {loadOptions.LoadFormat}");
                    Console.WriteLine($" - Number of worksheets: {workbook.Worksheets.Count}");
                    Console.WriteLine($" - First worksheet name: {workbook.Worksheets[0].Name}");

                    string outputPath = Path.ChangeExtension(filePath, ".converted.xlsx");
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($" - Saved as: {outputPath}");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load '{Path.GetFileName(filePath)}': {ex.Message}");
            }
        }

        private static LoadOptions GetLoadOptions(LoadFormat format)
        {
            switch (format)
            {
                case LoadFormat.Csv:
                    return new TxtLoadOptions(LoadFormat.Csv);
                case LoadFormat.Tsv:
                    return new TxtLoadOptions(LoadFormat.Tsv);
                case LoadFormat.Ods:
                    return new LoadOptions(LoadFormat.Ods);
                case LoadFormat.SpreadsheetML:
                    return new LoadOptions(LoadFormat.SpreadsheetML);
                default:
                    return new LoadOptions();
            }
        }
    }
}