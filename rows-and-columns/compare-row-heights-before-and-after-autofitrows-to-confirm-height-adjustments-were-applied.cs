// Title: Compare row height before and after AutoFitRows with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds long wrapped text to cell A1, records the initial row height, calls worksheet.AutoFitRows(), records the new height, evaluates the difference, and saves the workbook.
// Keywords: Aspose.Cells | AutoFitRows | row height | C# | .NET | measure row height | Excel row auto fit | worksheet.AutoFitRows
// Common Searches: Aspose.Cells AutoFitRows row height check | C# get row height before autofit | verify AutoFitRows effect on row height
// Developer Intent: Confirm that calling worksheet.AutoFitRows() modifies a row's height by comparing the measured values before and after the operation.
// Use Cases: Validate that text wrapping triggers a height increase after AutoFitRows for dynamic report generation. | Log before/after row heights to audit layout changes when programmatically creating Excel files. | Include an assertion in unit tests to ensure AutoFitRows adjusts row height as expected.
// AI Prompts: Generate C# code that records a row's height, runs worksheet.AutoFitRows(), then compares the two values to confirm a change using Aspose.Cells. | Explain how to detect programmatically whether AutoFitRows altered a specific row's height and recommend an appropriate tolerance for floating‑point comparison. | Provide a unit‑test snippet that asserts the row height increases after applying AutoFitRows to a wrapped‑text cell in Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that creates a workbook, adds long wrapped text to cell A1, records the initial row height, calls worksheet.AutoFitRows(), records the new height, evaluates the difference, and saves the workbook.
class CompareRowHeights
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add long text to cell A1 and enable text wrapping to force a height change
        cells["A1"].PutValue("This is a very long text that should cause the row height to increase after autofit.");
        Style style = cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        cells["A1"].SetStyle(style);

        // Record the row height before AutoFitRows
        double heightBefore = cells.GetRowHeight(0);
        Console.WriteLine($"Row 0 height before AutoFitRows: {heightBefore}");

        // Auto-fit all rows in the worksheet (auto‑fit rule)
        worksheet.AutoFitRows();

        // Record the row height after AutoFitRows
        double heightAfter = cells.GetRowHeight(0);
        Console.WriteLine($"Row 0 height after AutoFitRows: {heightAfter}");

        // Compare heights to confirm adjustment
        bool heightChanged = Math.Abs(heightAfter - heightBefore) > 0.01;
        Console.WriteLine($"Height adjusted: {heightChanged}");

        // Save the workbook (save rule)
        workbook.Save("RowHeightComparison.xlsx");
    }
}
