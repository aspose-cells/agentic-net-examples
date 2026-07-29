// Title: C# – Delete Blank Rows from the First Worksheet with Cells.DeleteBlankRows (default options) in Aspose.Cells
// Description: This C# sample builds a workbook, writes values to column A while leaving rows 3 and 5 empty, prints the sheet before and after cleanup, invokes Cells.DeleteBlankRows() using the built‑in settings (no DeleteOptions and no UpdateReference), and writes the outcome to DeletedBlankRows.xlsx.
// Keywords: Aspose.Cells | C# delete empty rows | Cells.DeleteBlankRows | remove blank rows .NET | default delete options | skip UpdateReference | worksheet row cleanup | optimize workbook size | Aspose.Cells API example | row deletion tutorial
// Common Searches: Aspose.Cells delete empty rows C# | Cells.DeleteBlankRows default behavior | how to remove blank rows without affecting formulas Aspose | C# delete rows in first worksheet Aspose.Cells | skip UpdateReference when deleting rows Aspose
// Developer Intent: Remove every empty row from the first sheet by calling the standard DeleteBlankRows method while preserving existing references.
// Use Cases: Sanitize imported CSV data by stripping placeholder rows before generating a report. | Prepare a template workbook for distribution by clearing unused rows left blank during data entry. | Decrease file size and boost performance after programmatically populating a sheet.
// AI Prompts: Show how to delete blank rows from a specific worksheet while keeping all formulas intact using Aspose.Cells for .NET. | Provide a C# snippet that removes empty rows only within a given range and updates named ranges accordingly. | Explain the effect of DeleteOptions.UpdateReference on formula references when rows are deleted in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# sample builds a workbook, writes values to column A while leaving rows 3 and 5 empty, prints the sheet before and after cleanup, invokes Cells.DeleteBlankRows() using the built‑in settings (no DeleteOptions and no UpdateReference), and writes the outcome to DeletedBlankRows.xlsx.
    public class DeleteBlankRowsDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the worksheet with data and intentional blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Row1");
            // Row 3 will be blank
            cells["A4"].PutValue("Row2");
            // Row 5 will be blank
            cells["A6"].PutValue("Row3");

            // Display rows before deletion (for demonstration)
            Console.WriteLine("Before deleting blank rows:");
            for (int i = 0; i <= 6; i++)
            {
                Console.WriteLine($"Row {i + 1}: '{cells[i, 0].StringValue}'");
            }

            // Delete all blank rows using the default method (no DeleteOptions, no UpdateReference)
            cells.DeleteBlankRows();

            // Display rows after deletion (for demonstration)
            Console.WriteLine("\nAfter deleting blank rows:");
            int maxRow = cells.MaxDataRow;
            for (int i = 0; i <= maxRow; i++)
            {
                Console.WriteLine($"Row {i + 1}: '{cells[i, 0].StringValue}'");
            }

            // Save the workbook to a file
            workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
        }
    }
}
