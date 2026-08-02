// Title: C# – Set Custom Worksheet Paper Size in Aspose.Cells Using Millimeter Dimensions
// Description: Shows a C# helper that takes width and height in millimeters, converts them to points (1 pt = 25.4/72 mm) and then to inches, sets PageSetup.PaperSize to Custom, and applies the size with PageSetup.CustomPaperSize on the first worksheet before saving the workbook.
// Keywords: Aspose.Cells | C# custom paper size | millimeter to point conversion | PageSetup.CustomPaperSize | Excel worksheet page setup | convert mm to inches | non‑standard page size | label sheet printing | PDF report page size | Aspose.Cells API
// Common Searches: Aspose.Cells set custom paper size C# | convert millimeters to points for Excel page setup | PageSetup.CustomPaperSize expects inches | how to define custom worksheet size in Aspose.Cells | C# convert mm to points Excel | custom paper size for PDF export Aspose.Cells
// Developer Intent: Create a reusable C# method that receives a Workbook and paper dimensions in millimeters, converts the values to the units required by Aspose.Cells, and applies a custom page size to the first worksheet.
// Use Cases: Generate a PDF report with a non‑standard page size (e.g., 100 mm × 150 mm) by setting a custom paper size before exporting. | Print label sheets or forms that require exact dimensions defined in millimeters. | Prepare a workbook for a specialized layout (e.g., custom brochures) where the default paper sizes are insufficient.
// AI Prompts: Write a C# method that accepts width and height in millimeters, converts them to points and inches, and sets a custom paper size on an Aspose.Cells worksheet. | Explain the step‑by‑step conversion from millimeters to points and then to inches needed for PageSetup.CustomPaperSize in Aspose.Cells. | Add validation and error handling to the custom paper size helper to manage negative or zero dimensions. | Show how to apply the custom paper size to all worksheets in a workbook instead of only the first one.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows a C# helper that takes width and height in millimeters, converts them to points (1 pt = 25.4/72 mm) and then to inches, sets PageSetup.PaperSize to Custom, and applies the size with PageSetup.CustomPaperSize on the first worksheet before saving the workbook.
    public static class CustomPaperSizeHelper
    {
        /// <param name="workbook">The workbook whose first worksheet will be modified.</param>
        /// <param name="widthMm">Paper width in millimeters.</param>
        /// <param name="heightMm">Paper height in millimeters.</param>
        public static void SetCustomPaperSizeInPoints(Workbook workbook, double widthMm, double heightMm)
        {
            // Convert millimeters to points.
            // 1 point = 1/72 inch, 1 inch = 25.4 mm  => 1 point = 25.4 / 72 mm
            const double mmPerPoint = 25.4 / 72.0;
            double widthPoints = widthMm / mmPerPoint;
            double heightPoints = heightMm / mmPerPoint;

            // Convert points to inches for the CustomPaperSize method.
            double widthInches = widthPoints / 72.0;
            double heightInches = heightPoints / 72.0;

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            PageSetup pageSetup = sheet.PageSetup;

            // Indicate that we are using a custom paper size.
            pageSetup.PaperSize = PaperSizeType.Custom;

            // Apply the custom size (method expects inches).
            pageSetup.CustomPaperSize(widthInches, heightInches);
        }

        // Example usage.
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create).
            Workbook workbook = new Workbook();

            // Set custom paper size to 100 mm x 150 mm.
            SetCustomPaperSizeInPoints(workbook, 100.0, 150.0);

            // Save the workbook (lifecycle rule: save).
            workbook.Save("CustomPaperSizeInPoints.xlsx");
        }
    }

    // Entry point for the console application.
    public static class Program
    {
        public static void Main()
        {
            try
            {
                CustomPaperSizeHelper.Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
