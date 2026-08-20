// Title: Toggle Worksheet View (Normal ↔ Page Break Preview) with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, access its first worksheet, and use the IsPageBreakPreview property to switch between Normal view and Page Break Preview based on a Boolean flag. The example also sets a standard zoom level, logs the current view state, and saves the file with a name that reflects the selected mode.
// Keywords: Aspose.Cells | C# | IsPageBreakPreview | worksheet view mode | page break preview | normal view | set worksheet zoom | save workbook programmatically | toggle view Aspose.Cells
// Common Searches: Aspose.Cells enable page break preview C# | toggle worksheet view normal page break preview Aspose | set IsPageBreakPreview property programmatically | save workbook with specific view mode using Aspose.Cells
// Developer Intent: Programmatically change a worksheet’s display mode between Normal view and Page Break Preview according to a user‑provided flag.
// Use Cases: Create a printable report that is saved in Page Break Preview to show exact pagination. | Provide a data‑entry workbook that defaults to Normal view for editing and exports a preview version for review. | Add a UI option in a WinForms or web app that lets users select the view mode before exporting the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that sets IsPageBreakPreview to true when a boolean variable is true and saves the file as 'ReportPreview.xlsx'. | Show an example that toggles the worksheet view based on user input, includes a zoom setting of 100%, and implements proper exception handling. | Explain how to switch between Normal view and Page Break Preview for all worksheets in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, access its first worksheet, and use the IsPageBreakPreview property to switch between Normal view and Page Break Preview based on a Boolean flag. The example also sets a standard zoom level, logs the current view state, and saves the file with a name that reflects the selected mode.
    public class ToggleWorksheetView
    {
        // Toggles the worksheet view based on the provided flag.
        // If enablePageBreakPreview is true, the sheet is shown in Page Break Preview mode;
        // otherwise it remains in Normal view.
        public static void Run(bool enablePageBreakPreview)
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Apply the view toggle (feature: IsPageBreakPreview property)
                worksheet.IsPageBreakPreview = enablePageBreakPreview;

                // Set zoom to a standard level for demonstration
                worksheet.Zoom = 100;

                // Output the current view state
                Console.WriteLine("IsPageBreakPreview: " + worksheet.IsPageBreakPreview);

                // Save the workbook (lifecycle: save)
                string fileName = enablePageBreakPreview ? "PageBreakPreview.xlsx" : "NormalView.xlsx";
                workbook.Save(fileName);
                Console.WriteLine($"Workbook saved as {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ToggleWorksheetView.Run: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                // Demonstrate both view modes
                ToggleWorksheetView.Run(true);
                ToggleWorksheetView.Run(false);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
