using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ToggleWorksheetVisibility
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the existing workbook
                string inputPath = "input.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and set visibility
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool shouldBeVisible = !sheet.Name.StartsWith("Hidden", StringComparison.OrdinalIgnoreCase);
                    sheet.IsVisible = shouldBeVisible;
                }

                // Path for the updated workbook
                string outputPath = "output.xlsx";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}