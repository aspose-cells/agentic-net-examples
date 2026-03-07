using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TabExcelToJson
    {
        public static void Run()
        {
            // Create a temporary TAB‑delimited file with sample data
            string sourcePath = Path.Combine(Path.GetTempPath(), "input_tab.txt");
            File.WriteAllText(sourcePath, "Name\tAge\tScore\nAlice\t30\t85.5\nBob\t25\t92");

            // Configure load options for a TAB‑delimited text file
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = '\t',          // Tab character as column separator
                ConvertNumericData = true // Convert numeric strings to numbers
            };

            // Load the workbook using the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true, // Export as JSON object even for a single sheet
                ExportEmptyCells = true,         // Include empty cells in the output
                HasHeaderRow = true              // Treat the first row as header
            };

            // Save the workbook as a JSON file
            string outputPath = Path.Combine(Path.GetTempPath(), "output.json");
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TabExcelToJson.Run();
        }
    }
}