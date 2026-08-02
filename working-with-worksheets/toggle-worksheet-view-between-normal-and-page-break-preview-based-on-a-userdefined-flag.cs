// Title: Toggle worksheet view (Normal vs Page Break Preview) with Aspose.Cells for .NET
// Description: Creates a workbook, accesses the first worksheet, and uses a Boolean flag to set either IsPageBreakPreview or ViewType to NormalView or PageBreakPreview, then saves the file.
// Keywords: Aspose.Cells worksheet view | C# IsPageBreakPreview | PageBreakPreview view | NormalView Aspose | toggle worksheet view | set ViewType programmatically | Aspose.Cells .NET example
// Common Searches: Aspose.Cells change worksheet to Page Break Preview C# | How to switch between Normal view and Page Break Preview in Aspose.Cells | Set worksheet view type with a boolean flag Aspose.Cells | C# toggle worksheet view Aspose.Cells example | IsPageBreakPreview property usage Aspose.Cells
// Developer Intent: Programmatically set a worksheet’s display mode to Normal view or Page Break Preview based on a runtime flag.
// Use Cases: Automatically open generated reports in Page Break Preview to show print layout. | Provide a UI toggle that lets end‑users choose Normal or Page Break Preview before exporting. | Apply different view modes to multiple worksheets in a single workbook according to layout needs.
// AI Prompts: Generate C# code that reads a config value and switches an Aspose.Cells worksheet between NormalView and PageBreakPreview. | Show how to revert a worksheet’s ViewType from PageBreakPreview back to NormalView after saving. | Explain the difference between IsPageBreakPreview and ViewType and when to use each for dynamic view control.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, accesses the first worksheet, and uses a Boolean flag to set either IsPageBreakPreview or ViewType to NormalView or PageBreakPreview, then saves the file.
    public class ToggleWorksheetView
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // User‑defined flag: true for Page Break Preview, false for Normal view
            bool showPageBreakPreview = true; // change as needed

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Toggle the view based on the flag
            worksheet.IsPageBreakPreview = showPageBreakPreview;

            // Optionally, you can also set ViewType directly:
            // worksheet.ViewType = showPageBreakPreview ? ViewType.PageBreakPreview : ViewType.NormalView;

            // Save the workbook (lifecycle rule: save)
            string outputPath = "ToggleWorksheetView_Output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
