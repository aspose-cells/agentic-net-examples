using System;
using Aspose.Cells;

namespace AsposeCellsJsonConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLS workbook
            string sourcePath = "input.xls";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(sourcePath);

            // Create default JSON save options (all settings are default)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file using the default conversion settings
            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{outputPath}'.");
        }
    }
}