using System;
using System.IO;
using Aspose.Cells;
using Range = Aspose.Cells.Range;

namespace BatchExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFolder = args.Length > 0 ? args[0] : @"C:\InputExcelFiles";
            string outputFolder = args.Length > 1 ? args[1] : @"C:\OutputJsonFiles";

            Directory.CreateDirectory(outputFolder);

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder '{inputFolder}' does not exist.");
                return;
            }

            foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xls*"))
            {
                try
                {
                    Workbook workbook = new Workbook(excelPath);

                    JsonSaveOptions jsonOptions = new JsonSaveOptions
                    {
                        AlwaysExportAsJsonObject = true,
                        SkipEmptyRows = true,
                        ExportAsString = true,
                        Indent = "  "
                    };

                    string jsonFileName = Path.Combine(
                        outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + ".json");

                    workbook.Save(jsonFileName, jsonOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(excelPath)}' to JSON successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(excelPath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}