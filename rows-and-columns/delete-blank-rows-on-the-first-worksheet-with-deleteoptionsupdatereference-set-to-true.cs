// Title: Delete blank rows on the first worksheet with UpdateReference in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert data with intentional empty rows, configure DeleteOptions.UpdateReference = true, remove all blank rows from the first worksheet, and save the file so that formulas and named ranges automatically adjust to the new row positions.
// Keywords: Aspose.Cells delete blank rows | DeleteBlankRows UpdateReference | C# remove empty rows Aspose.Cells | UpdateReference after row deletion | Aspose.Cells DeleteOptions example
// Common Searches: Aspose.Cells delete empty rows keep formulas | DeleteBlankRows with UpdateReference true C# | Remove blank rows from first worksheet Aspose.Cells | How to preserve cell references when deleting rows in Aspose.Cells
// Developer Intent: Remove every empty row from the first worksheet while automatically updating all dependent formulas and references.
// Use Cases: Clean up generated reports by eliminating placeholder rows without breaking chart data ranges. | Prepare export templates by deleting blank rows while retaining correct formula calculations. | Pre‑process imported data that contains sporadic empty rows, ensuring named ranges and references stay intact.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete blank rows on the first worksheet with DeleteOptions.UpdateReference set to true. | Explain the effect of DeleteOptions.UpdateReference on formulas after calling Cells.DeleteBlankRows in Aspose.Cells. | Provide a step‑by‑step tutorial for removing empty rows from a worksheet while preserving named ranges and formulas using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert data with intentional empty rows, configure DeleteOptions.UpdateReference = true, remove all blank rows from the first worksheet, and save the file so that formulas and named ranges automatically adjust to the new row positions.
    public class DeleteBlankRowsWithUpdateReference
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add sample data with blank rows
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue("Row1");
                // Row 3 is left blank intentionally
                cells["A4"].PutValue("Row2");
                // Row 5 is left blank intentionally
                cells["A6"].PutValue("Row3");

                // Configure DeleteOptions to update references after deletion
                DeleteOptions options = new DeleteOptions
                {
                    UpdateReference = true
                };

                // Delete all blank rows on the first worksheet using the options
                cells.DeleteBlankRows(options);

                // Save the modified workbook
                workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteBlankRowsWithUpdateReference.Run();
        }
    }
}
