// Title: Unfreeze All Panes in an Aspose.Cells Worksheet (C#)
// Description: This example creates a workbook, optionally freezes panes, then removes any frozen rows or columns using the Worksheet.UnFreezePanes method (or FreezePanes(0,0)). The file is saved as UnfreezePanesDemo.xlsx, demonstrating how to reset pane settings in Excel with Aspose.Cells for .NET.
// Keywords: Aspose.Cells UnFreezePanes C# | unfreeze Excel panes .NET | clear frozen rows Aspose.Cells | Worksheet.UnFreezePanes example | FreezePanes(0,0) Aspose.Cells | reset pane freezing C# | Aspose.Cells worksheet freeze/unfreeze
// Common Searches: how to unfreeze panes using Aspose.Cells C# | remove frozen rows and columns Aspose.Cells .NET | Worksheet.UnFreezePanes method example | unfreeze all panes in Excel with Aspose.Cells | FreezePanes(0,0) to clear pane freeze
// Developer Intent: Programmatically clear any frozen rows or columns so the worksheet displays with all panes unfrozen.
// Use Cases: Generate a report that temporarily freezes header rows for preview, then unfreeze before delivering the final file. | Create a utility that toggles pane freezing based on user preferences, requiring a clean reset before applying new settings. | Automate workbook cleanup in a CI pipeline to ensure no frozen panes remain in exported Excel files.
// AI Prompts: Write C# code using Aspose.Cells to unfreeze all panes in a worksheet, including a verification step that checks the freeze state before and after. | Explain the difference between Worksheet.UnFreezePanes() and Worksheet.FreezePanes(0,0) in Aspose.Cells, and show code examples for each. | Provide a concise tutorial that demonstrates freezing specific rows/columns, then removing the freeze with UnFreezePanes, highlighting best practices for Excel automation.

using System;
using Aspose.Cells;

// This example creates a workbook, optionally freezes panes, then removes any frozen rows or columns using the Worksheet.UnFreezePanes method (or FreezePanes(0,0)). The file is saved as UnfreezePanesDemo.xlsx, demonstrating how to reset pane settings in Excel with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Freeze some panes to demonstrate the unfreeze operation
        worksheet.FreezePanes(2, 2, 2, 2);

        // Unfreeze all panes in the worksheet
        worksheet.UnFreezePanes();

        // Save the workbook to verify that panes are unfrozen
        workbook.Save("UnfreezePanesDemo.xlsx");
    }
}
