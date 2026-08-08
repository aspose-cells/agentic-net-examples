// Title: C# – Delete Blank Rows and Columns in Aspose.Cells while Keeping Formulas Intact (UpdateReference = false)
// Description: Demonstrates how to create or load a workbook, set DeleteOptions.UpdateReference to false, and call DeleteBlankRows and DeleteBlankColumns on the first worksheet. The operation removes empty rows and columns without altering formula references, then saves the file as output.xlsx.
// Keywords: Aspose.Cells C# delete blank rows | Aspose.Cells DeleteBlankColumns | UpdateReference false | preserve Excel formulas | remove empty rows Aspose.Cells | Excel cleanup .NET | DeleteOptions Aspose.Cells | C# Excel automation | Aspose.Cells workbook trimming
// Common Searches: Aspose.Cells delete empty rows without changing formulas | C# DeleteBlankColumns UpdateReference example | How to keep formula references when removing blank rows in Aspose.Cells | Remove blank rows and columns from first worksheet Aspose.Cells .NET | DeleteOptions.UpdateReference usage in Excel automation
// Developer Intent: Remove all blank rows and columns from the first worksheet while ensuring existing formulas remain unchanged.
// Use Cases: Sanitize user‑generated spreadsheets by stripping out blank rows/columns before further processing. | Prepare a report template where placeholder rows are eliminated without breaking dependent calculations. | Automate data‑export cleanup for downstream systems while retaining all formula logic.
// AI Prompts: Write C# code using Aspose.Cells to delete blank rows and columns on a specific worksheet without updating formula references. | Explain the effect of DeleteOptions.UpdateReference on formulas when deleting rows or columns in Aspose.Cells. | Provide a step‑by‑step guide to load an existing workbook, apply DeleteBlankRows/DeleteBlankColumns with UpdateReference disabled, and save the result.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create or load a workbook, set DeleteOptions.UpdateReference to false, and call DeleteBlankRows and DeleteBlankColumns on the first worksheet. The operation removes empty rows and columns without altering formula references, then saves the file as output.xlsx.
    public class DeleteBlankRowsAndColumnsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Example data with blank rows and columns and formulas
                worksheet.Cells["A1"].PutValue(10);
                worksheet.Cells["B1"].PutValue(20);
                worksheet.Cells["C1"].Formula = "=A1+B1"; // Formula referencing columns A and B
                worksheet.Cells["A3"].PutValue(30);      // Row 2 is blank

                // Set delete options with UpdateReference disabled to preserve formulas
                DeleteOptions deleteOptions = new DeleteOptions
                {
                    UpdateReference = false // Do not adjust formulas after deletion
                };

                // Delete blank rows and columns on the first worksheet
                worksheet.Cells.DeleteBlankRows(deleteOptions);
                worksheet.Cells.DeleteBlankColumns(deleteOptions);

                // Save the modified workbook
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
