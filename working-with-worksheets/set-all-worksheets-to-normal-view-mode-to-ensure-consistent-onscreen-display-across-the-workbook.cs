// Title: Aspose.Cells C# – Set Every Worksheet to Normal View Mode
// Description: Shows how to use Aspose.Cells for .NET to change the view type of all worksheets in a workbook to NormalView. The sample creates a workbook, adds extra sheets, loops through each worksheet, sets ViewType.NormalView, and saves the file, guaranteeing a consistent on‑screen display for every sheet.
// Keywords: Aspose.Cells | C# | Set worksheet view | NormalView | ViewType | Excel workbook view mode | iterate worksheets | programmatic Excel view | Aspose.Cells example | Excel UI view
// Common Searches: Aspose.Cells set all worksheets to Normal view | C# loop through worksheets change view type | How to apply NormalView to every sheet in Aspose.Cells | Set default view mode for Excel workbook using Aspose.Cells .NET | Change worksheet view programmatically Aspose.Cells
// Developer Intent: Apply Normal view to every worksheet in an Excel workbook with Aspose.Cells.
// Use Cases: Standardize the on‑screen appearance of generated reports before distribution. | Reset view settings after adding or removing sheets to keep a uniform layout. | Prepare a workbook for screen capture or preview by ensuring all sheets use Normal view.
// AI Prompts: Generate C# code that iterates through all worksheets in an Aspose.Cells workbook and sets each sheet's ViewType to NormalView, then saves the file. | Explain how to toggle between NormalView and PageLayout view for selected worksheets using Aspose.Cells for .NET. | Provide step‑by‑step instructions to enforce a consistent view mode on all sheets of an existing Excel file with Aspose.Cells.

using Aspose.Cells;
using System;

// Shows how to use Aspose.Cells for .NET to change the view type of all worksheets in a workbook to NormalView. The sample creates a workbook, adds extra sheets, loops through each worksheet, sets ViewType.NormalView, and saves the file, guaranteeing a consistent on‑screen display for every sheet.
class SetAllWorksheetsToNormalView
{
    static void Main()
    {
        // Create a new workbook (contains one default worksheet)
        Workbook workbook = new Workbook();

        // Add additional worksheets for demonstration purposes
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Iterate through all worksheets and set the view type to NormalView
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ViewType = ViewType.NormalView;
        }

        // Save the workbook with the updated view settings
        workbook.Save("AllSheetsNormalView.xlsx");
    }
}
