// Title: Iterate over a custom display range (B2:D5) and write each cell's address using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a CellArea for B2:D5, loops through every cell in that area, and writes the cell's address into the cell with Aspose.Cells. | Show how to set a visible range on a worksheet, fill the cells with sample data, and save the workbook as an .xlsx file using Aspose.Cells for .NET. | Provide a snippet that builds a workbook, ensures the output folder exists, enumerates only the cells inside a specified CellArea, and then saves the file.
// Common Searches: Aspose.Cells C# iterate only cells within a defined display range | how to set a custom visible area B2:D5 in Aspose.Cells and loop through it | write cell address into each cell of a specific range using Aspose.Cells .NET | save workbook after modifying a specific CellArea with Aspose.Cells
// Tags: CellArea loop Aspose.Cells C# | custom visible range definition Aspose.Cells | populate cells with address Aspose.Cells | export workbook to .xlsx Aspose.Cells | enumerate cells in specific area Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Creates a workbook, defines a CellArea covering B2:D5 as the visible range, iterates through those cells to write each cell's address, and saves the result as DisplayRangeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Define the visible (display) range B2:D5 using zero‑based indexes
            var visibleArea = new CellArea
            {
                StartRow = 1,      // Row 2 (zero‑based)
                StartColumn = 1,   // Column B (zero‑based)
                EndRow = 4,        // Row 5
                EndColumn = 3      // Column D
            };

            // Enumerate only the cells inside the defined range
            for (int row = visibleArea.StartRow; row <= visibleArea.EndRow; row++)
            {
                for (int col = visibleArea.StartColumn; col <= visibleArea.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Write the cell address into the cell as an example operation
                    cell.PutValue(cell.Name);
                }
            }

            // Define output file path
            string outputPath = "DisplayRangeDemo.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
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
