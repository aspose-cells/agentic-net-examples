using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ToggleWorksheetViewDemo
    {
        // Toggles the view of the first worksheet based on the flag.
        // If isPageBreakPreview is true, the sheet is shown in Page Break Preview mode;
        // otherwise it is shown in Normal view.
        public static void Run(bool isPageBreakPreview, string outputPath)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the view mode according to the flag
            worksheet.IsPageBreakPreview = isPageBreakPreview;

            // Optional: set zoom to a comfortable level for demonstration
            worksheet.Zoom = 100;

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            // Toggle to Page Break Preview and save
            Run(true, "PageBreakPreviewDemo.xlsx");

            // Toggle back to Normal view and save
            Run(false, "NormalViewDemo.xlsx");

            Console.WriteLine("Worksheets saved with respective view settings.");
        }
    }
}