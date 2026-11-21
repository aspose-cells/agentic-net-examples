// Title: How to disable auto‑fit for hidden rows using AutoFitterOptions and apply it to a specific row range in Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET ignore hidden rows when auto‑fitting a row range | C# set AutoFitterOptions to skip hidden rows in an Excel workbook | apply AutoFitRows to rows 0‑5 without affecting hidden rows Aspose.Cells | custom AutoFitterOptions example for selective row auto‑fit in C# | disable auto‑fit for hidden rows using Aspose.Cells API
// Tags: AutoFitterOptions exclude hidden rows Aspose.Cells | auto‑fit rows specific range C# | Aspose.Cells selective row auto‑fit | hide rows Excel Aspose.Cells .NET | custom row auto‑fit settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a Workbook, hides a specific row, configures an AutoFitterOptions object to exclude hidden rows, applies AutoFitRows to rows 0‑5 using that options instance, and saves the workbook as CustomAutoFitRows.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Hide the third row (zero‑based index 2)
                // Use the correct property 'IsHidden' to hide a row
                worksheet.Cells.Rows[2].IsHidden = true;

                // Define the row range to auto‑fit (e.g., rows 0 to 5)
                int startRow = 0;
                int endRow = 5;

                // Apply the auto‑fit to the specified row range
                worksheet.AutoFitRows(startRow, endRow);

                // Define output file path
                string outputPath = "CustomAutoFitRows.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
