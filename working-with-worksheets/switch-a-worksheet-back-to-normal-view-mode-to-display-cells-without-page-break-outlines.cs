// Title: Aspose.Cells for .NET: Change Worksheet from Page Break Preview to Normal View (C#)
// Description: Shows how to programmatically switch a worksheet from Page Break Preview to Normal view with Aspose.Cells for .NET. The sample toggles the IsPageBreakPreview flag (or assigns ViewType.NormalView) to hide page‑break outlines and then saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | worksheet view | Normal view | Page Break Preview | IsPageBreakPreview | ViewType.NormalView | hide page break outlines | switch worksheet view mode | programmatic view change
// Common Searches: Aspose.Cells set worksheet normal view C# | disable page break preview Aspose.Cells .NET | hide page break outlines Aspose.Cells | change worksheet view type programmatically | Aspose.Cells ViewType.NormalView example
// Developer Intent: The developer wants to display a worksheet without page‑break outlines by converting it from Page Break Preview to Normal view in a .NET application.
// Use Cases: Generate reports that open directly in Normal view for end users. | Convert workbooks saved in Page Break Preview to a clean view before distribution. | Provide a UI toggle that lets users switch between preview and normal layouts on the fly.
// AI Prompts: Write C# code using Aspose.Cells to change a worksheet from Page Break Preview to Normal view and save the file. | Explain the difference between the IsPageBreakPreview property and setting ViewType to NormalView in Aspose.Cells. | Create a reusable method that accepts a Worksheet object and ensures it is displayed in Normal view, handling both IsPageBreakPreview and ViewType properties.

using System;
using Aspose.Cells;

// Shows how to programmatically switch a worksheet from Page Break Preview to Normal view with Aspose.Cells for .NET. The sample toggles the IsPageBreakPreview flag (or assigns ViewType.NormalView) to hide page‑break outlines and then saves the workbook.
class SwitchToNormalView
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Set to Page Break Preview first to demonstrate the switch
        worksheet.IsPageBreakPreview = true;

        // Switch the worksheet back to Normal view mode
        // This hides page break outlines and shows cells normally
        worksheet.IsPageBreakPreview = false;
        // Alternatively you can use:
        // worksheet.ViewType = ViewType.NormalView;

        // Save the workbook
        workbook.Save("NormalViewDemo.xlsx");
    }
}
