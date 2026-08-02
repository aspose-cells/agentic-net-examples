// Title: Reorder worksheets: Move a sheet to the first tab and set it active with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add multiple sheets, relocate a chosen worksheet (e.g., Sheet3) to index 0, activate it, and save the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | MoveTo method | worksheet reorder | first tab | set active sheet | Excel automation | workbook manipulation | reorder worksheets programmatically
// Common Searches: Aspose.Cells move worksheet to first position C# | How to set active sheet after reordering Aspose.Cells | MoveTo index 0 example Aspose.Cells | Reorder Excel tabs with Aspose.Cells .NET | Change worksheet tab order programmatically
// Developer Intent: Programmatically move a specific worksheet to the first position in a workbook and make it the active sheet before saving.
// Use Cases: Place a dashboard or summary sheet at the front of an automatically generated report. | Prioritize a newly created analysis worksheet in a batch‑processing workflow. | Ensure the default view opens on a designated sheet after dynamic tab reordering.
// AI Prompts: Write C# code using Aspose.Cells to move a worksheet named "Report" to the first tab and set it as the active sheet before saving. | Explain the behavior of the Worksheet.MoveTo method when given index values, including error handling for out‑of‑range indices. | Provide a complete example that clears default sheets, adds several worksheets, reorders them, activates a specific sheet, and saves the workbook as XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetReorder
{
    // Shows how to create a workbook, add multiple sheets, relocate a chosen worksheet (e.g., Sheet3) to index 0, activate it, and save the result using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Remove the default worksheet to avoid duplicate names
                workbook.Worksheets.Clear();

                // Add sample worksheets with unique names
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Retrieve the worksheet you want to move (e.g., "Sheet3")
                Worksheet sheetToMove = workbook.Worksheets["Sheet3"];

                // Relocate the worksheet to the first position (index 0)
                // This makes it the leftmost tab when the workbook is opened
                sheetToMove.MoveTo(0);

                // Optionally set it as the active sheet so it is displayed first
                workbook.Worksheets.ActiveSheetIndex = sheetToMove.Index;

                // Save the workbook (lifecycle: save)
                workbook.Save("ReorderedWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
