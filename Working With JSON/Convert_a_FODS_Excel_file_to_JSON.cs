using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FodsToJsonConverter
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source FODS file
            string sourcePath = "input.fods";

            // Desired output JSON file path
            string outputPath = "output.json";

            // Load the FODS file using LoadOptions with the appropriate format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook as JSON
            workbook.Save(outputPath, SaveFormat.Json);

            Console.WriteLine($"FODS file '{sourcePath}' has been successfully converted to JSON at '{outputPath}'.");
        }
    }
}