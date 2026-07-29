// Title: Remove blank rows with DeleteOptions.UpdateReference in AspNet Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, insert data with intentional empty rows, configure DeleteOptions.UpdateReference = true, call cells.DeleteBlankRows(options) to purge blank rows from the first worksheet, and save the file while automatically updating formulas and cell references.
// Keywords: Aspose.Cells | DeleteBlankRows | DeleteOptions | UpdateReference | C# | .NET | remove empty rows | adjust formulas | worksheet cleanup | Excel automation
// Common Searches: Aspose.Cells delete blank rows C# | DeleteOptions.UpdateReference example | How to remove empty rows while preserving formulas Aspose.Cells | cells.DeleteBlankRows with options .NET | Delete blank rows first worksheet Aspose.Cells
// Developer Intent: Delete every empty row on the first worksheet and automatically refresh any dependent cell references or formulas.
// Use Cases: Clean generated reports by eliminating placeholder rows before distribution. | Prepare data extracts where blank rows would break downstream processing. | Automate workbook sanitization to keep formula integrity after programmatic row insertion. | Maintain chart data ranges by removing gaps without breaking references.
// AI Prompts: Write C# code using Aspose.Cells to delete blank rows on a specific worksheet with UpdateReference enabled. | Explain the impact of DeleteOptions.UpdateReference on formulas and named ranges after row deletion in Aspose.Cells. | Show how to modify the sample to target a named worksheet instead of the first sheet. | Provide a step‑by‑step guide for deleting empty rows while preserving chart series in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert data with intentional empty rows, configure DeleteOptions.UpdateReference = true, call cells.DeleteBlankRows(options) to purge blank rows from the first worksheet, and save the file while automatically updating formulas and cell references.
    public class DeleteBlankRowsWithUpdateReference
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate the worksheet with data and intentional blank rows
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue("Row 1");
                // Row 3 will be blank
                cells["A4"].PutValue("Row 2");
                // Row 5 will be blank
                cells["A6"].PutValue("Row 3");

                // Set up DeleteOptions with UpdateReference = true
                DeleteOptions options = new DeleteOptions
                {
                    UpdateReference = true
                };

                // Delete all blank rows on the first worksheet using the options
                cells.DeleteBlankRows(options);

                // Save the modified workbook
                workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully as DeletedBlankRows.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
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
