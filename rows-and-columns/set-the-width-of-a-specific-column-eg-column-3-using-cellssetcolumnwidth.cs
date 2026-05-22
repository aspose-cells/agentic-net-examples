using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetColumnWidthExample
    {
        // Entry point required by the .NET runtime
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (no template file needed)
            Workbook workbook = new Workbook();

            // Get the Cells collection of the first worksheet
            Cells cells = workbook.Worksheets[0].Cells;

            // Set the width of column 3 (zero‑based index 2) to 25.5 characters
            cells.SetColumnWidth(2, 25.5);

            // Add sample data to demonstrate the column width
            cells["C1"].PutValue("Column 3 width set to 25.5 characters");

            // Define output file path
            string outputPath = "Column3WidthDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}