using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SuppressPrintErrorsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Suppress error values during printing by displaying them as blank
            sheet.PageSetup.PrintErrors = PrintErrorsType.PrintErrorsBlank;

            // Define output file path
            string outputPath = "SuppressPrintErrors.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}