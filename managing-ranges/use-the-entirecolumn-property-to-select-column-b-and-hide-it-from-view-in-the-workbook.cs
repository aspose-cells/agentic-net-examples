// Title: Hide column B in an Excel workbook using Aspose.Cells .NET EntireColumn property
// AI Prompts: Write C# code that selects column B via the EntireColumn property and sets its IsHidden flag to true with Aspose.Cells. | Show how to toggle the visibility of a specific column (e.g., column B) in an Excel file using Aspose.Cells .NET APIs. | Provide an example that hides multiple columns (such as C and D) by iterating over their EntireColumn ranges in Aspose.Cells.
// Common Searches: Aspose.Cells .NET hide column B using EntireColumn property | C# programmatically hide a specific Excel column with Aspose.Cells | How to set column visibility in an Aspose.Cells workbook | Using EntireColumn to hide columns in Aspose.Cells C# example | Save Excel file after hiding columns with Aspose.Cells
// Tags: Aspose.Cells hide column using EntireColumn | C# set column IsHidden Aspose.Cells | Aspose.Cells column visibility .NET | Excel workbook hide column programmatically | Aspose.Cells save workbook after column hide

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a new workbook, accesses the first worksheet, selects column B via the Columns collection, hides it by setting IsHidden = true, ensures the output directory exists, saves the workbook as output.xlsx, and handles any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Hide column B (zero‑based index 1) using the Column object's IsHidden property
                var column = worksheet.Cells.Columns[1];
                column.IsHidden = true;
                // Alternatively: worksheet.Cells.HideColumn(1);

                // Output file path
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
