// Title: How to set the height of row 5 to 30 points in an Excel file using Aspose.Cells SetRowHeight in C#
// Common Searches: C# Aspose.Cells set row height to specific points example | Set height of row 5 in .xlsx using Aspose.Cells Cells.SetRowHeight | Programmatically change Excel row height in C# with Aspose.Cells | Adjust row dimensions in a workbook created with Aspose.Cells C# | How to use Cells.SetRowHeight to define row size in Aspose.Cells
// Tags: Aspose.Cells Cells.SetRowHeight usage | C# set Excel row height points | modify specific row height Aspose.Cells | row height adjustment .xlsx C# | programmatic row dimension control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, accesses the first Worksheet, and calls sheet.Cells.SetRowHeight(4, 30) to set row 5 (zero‑based index 4) to a height of 30 points, then saves the workbook as output.xlsx while handling any exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set the height of row 5 (zero‑based index 4) to 30 points
                sheet.Cells.SetRowHeight(4, 30);

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
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
