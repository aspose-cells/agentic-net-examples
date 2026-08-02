// Title: C# – Set Worksheet Zoom (150%) and Freeze Top 4 Rows with Aspose.Cells
// Description: Creates a new workbook, sets the first worksheet to Normal view, applies a 150 % zoom, freezes the first four rows while leaving columns unfrozen, and saves the file as CustomViewFreezeRows.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | set worksheet zoom | freeze rows | FreezePanes | worksheet view type | NormalView | custom zoom level | Excel automation .NET | freeze top rows
// Common Searches: Aspose.Cells set zoom level C# | Freeze first rows Aspose.Cells .NET | How to freeze rows without columns in Aspose.Cells | Custom worksheet view example Aspose.Cells | Set NormalView and zoom in Aspose.Cells
// Developer Intent: Create a workbook, apply a 150 % zoom, and lock the first four rows on the first worksheet.
// Use Cases: Display header rows constantly while users scroll through large data sets, improving readability with a larger zoom. | Prepare a reporting template that opens at a predefined zoom for consistent screen layout across devices. | Build an Excel‑based data‑entry form where the top rows stay fixed, ensuring users always see key fields.
// AI Prompts: Show C# code to set a 200 % zoom and freeze the top three rows using Aspose.Cells. | Give an Aspose.Cells example that applies NormalView, sets a custom zoom, and freezes both rows and columns simultaneously. | Explain the FreezePanes parameters for freezing only rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new workbook, sets the first worksheet to Normal view, applies a 150 % zoom, freezes the first four rows while leaving columns unfrozen, and saves the file as CustomViewFreezeRows.xlsx using Aspose.Cells for .NET.
class SetCustomViewAndFreezeRows
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the view type (optional, default is NormalView)
        sheet.ViewType = ViewType.NormalView;

        // Set a custom zoom level (e.g., 150%)
        sheet.Zoom = 150;

        // Freeze the first 4 rows (row index 4) while keeping columns unfrozen
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        sheet.FreezePanes(4, 0, 4, 0);

        // Save the workbook to a file
        workbook.Save("CustomViewFreezeRows.xlsx");
    }
}
