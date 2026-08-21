// Title: Unfreeze all worksheet panes in Aspose.Cells for .NET with FreezePanes(0,0,0,0)
// Description: This C# example creates a workbook, freezes panes at row 3/column 3, then removes every frozen pane by calling FreezePanes(0,0,0,0) and saves the result as UnfreezePanesByUnfreezePanes.xlsx.
// Keywords: Aspose.Cells unfreeze panes | FreezePanes zero parameters | clear frozen rows C# | reset worksheet freeze .NET | remove pane freeze Aspose
// Common Searches: how to unfreeze panes Aspose.Cells | reset FreezePanes to zero C# | clear frozen rows and columns Aspose.Cells .NET | unfreeze all worksheet panes programmatically
// Developer Intent: Remove any frozen rows or columns from a worksheet by resetting the FreezePanes method to zero values.
// Use Cases: Prepare a report that must open without locked sections by clearing previous pane freezes. | Refresh a generated template so that no panes remain frozen after layout changes. | Automate cleanup of user‑modified workbooks before distribution or archival.
// AI Prompts: Generate C# code that loads an existing Excel file with Aspose.Cells, detects frozen panes, and unfreezes them using FreezePanes(0,0,0,0). | Explain how to check the FreezePanes state of a worksheet before resetting it in Aspose.Cells for .NET. | Provide a robust try‑catch pattern that unfreezes panes and logs any errors during workbook saving.

using System;
using Aspose.Cells;

// This C# example creates a workbook, freezes panes at row 3/column 3, then removes every frozen pane by calling FreezePanes(0,0,0,0) and saves the result as UnfreezePanesByUnfreezePanes.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze some panes (example)
            worksheet.FreezePanes(3, 3, 3, 3);

            // Unfreeze all panes by resetting the freeze parameters
            worksheet.FreezePanes(0, 0, 0, 0);

            // Save the workbook
            string outputPath = "UnfreezePanesByUnfreezePanes.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
