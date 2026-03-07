using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MhtCsvToJsonConverter
    {
        public static void Run()
        {
            string sourcePath = "input.mht";

            Workbook workbook;

            if (File.Exists(sourcePath))
            {
                var loadOptions = new LoadOptions(LoadFormat.Csv);
                workbook = new Workbook(sourcePath, loadOptions);
            }
            else
            {
                // Create a simple CSV content as fallback
                string csvContent = "Name,Age\nJohn,30\nJane,25";
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(csvContent)))
                {
                    var loadOptions = new LoadOptions(LoadFormat.Csv);
                    workbook = new Workbook(ms, loadOptions);
                }
            }

            var jsonOptions = new JsonSaveOptions
            {
                ExportNestedStructure = true,
                SkipEmptyRows = true,
                AlwaysExportAsJsonObject = true
            };

            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MhtCsvToJsonConverter.Run();
        }
    }
}