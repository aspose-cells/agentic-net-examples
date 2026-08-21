// Title: Move a worksheet to the first tab and set it active with Aspose.Cells for .NET
// Description: Creates a workbook, adds several sheets, moves a chosen worksheet (e.g., Sheet3) to index 0, marks it as the active and selected sheet, and saves the workbook.
// Keywords: Aspose.Cells | move worksheet | first tab | set active sheet | reorder worksheets | .NET | C# | Workbook.MoveTo | ActiveSheetIndex | worksheet index 0
// Common Searches: Aspose.Cells move worksheet to first tab | set active sheet after moving worksheet Aspose.Cells | reorder worksheets programmatically .NET | how to make a sheet the first tab in Aspose.Cells | select worksheet on workbook open Aspose.Cells
// Developer Intent: Reorder the worksheets so a specific sheet becomes the first tab and is active when the workbook is opened.
// Use Cases: Place a dashboard sheet at the beginning of a generated report for immediate visibility. | Ensure a summary sheet opens as the active tab after programmatic sheet reordering. | Create a template where the most important worksheet is always positioned at index 0 and pre‑selected.
// AI Prompts: Generate C# code using Aspose.Cells to move a worksheet named "Report" to the first position and make it the active sheet. | Explain how to reorder multiple worksheets in an Aspose.Cells workbook while keeping the first sheet selected on open. | Provide error‑handling examples for moving a worksheet to index 0 when the target sheet might be missing.

using System;
using Aspose.Cells;

namespace WorksheetRelocationDemo
{
    // Creates a workbook, adds several sheets, moves a chosen worksheet (e.g., Sheet3) to index 0, marks it as the active and selected sheet, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains a default sheet)
                Workbook workbook = new Workbook();

                // Remove the default sheet to avoid duplicate name errors
                workbook.Worksheets.Clear();

                // Add sample worksheets with unique names
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Get the worksheet to be moved (e.g., "Sheet3")
                Worksheet sheetToMove = workbook.Worksheets["Sheet3"];

                // Move the worksheet to the first position (index 0)
                sheetToMove.MoveTo(0);

                // Make the moved sheet the active and selected sheet
                workbook.Worksheets.ActiveSheetIndex = sheetToMove.Index;
                sheetToMove.IsSelected = true;

                // Save the workbook
                workbook.Save("RelocatedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
