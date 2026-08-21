// Title: C# – Detect if a Worksheet’s Custom Paper Size Exceeds A3 Using Aspose.Cells
// Description: Provides a C# helper method that reads a worksheet’s PageSetup, confirms the PaperSize is set to Custom, and returns true when the PaperWidth or PaperHeight (in inches) is larger than the A3 limits (≈11.69 × 16.54 in). Includes sample code that sets custom sizes, runs the check, and saves the workbook.
// Keywords: Aspose.Cells | C# | custom paper size | A3 dimensions | Worksheet PageSetup | PaperWidth | PaperHeight | PaperSizeType.Custom | Print layout validation | CustomPaperSize method
// Common Searches: Aspose.Cells check custom paper size larger than A3 | C# function to compare worksheet page size with A3 | How to detect oversized custom paper in Aspose.Cells | Validate print area against A3 using Aspose.Cells for .NET | IsCustomPaperSizeExceedsA3 example
// Developer Intent: Determine programmatically whether a worksheet’s custom paper dimensions exceed the standard A3 size.
// Use Cases: Prevent printer errors by flagging worksheets whose custom size is larger than A3 before exporting to PDF. | Enforce corporate print‑size policies when generating automated reports with Aspose.Cells. | Log a warning or automatically adjust scaling/orientation when an oversized custom page is detected.
// AI Prompts: Generate unit tests for IsCustomPaperSizeExceedsA3 covering sizes below, equal to, and above A3. | Write a C# snippet that logs a warning and switches to portrait orientation if the custom paper size exceeds A3. | Create a PowerShell script that scans all worksheets in a workbook and reports any that exceed A3 dimensions using the provided helper.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Provides a C# helper method that reads a worksheet’s PageSetup, confirms the PaperSize is set to Custom, and returns true when the PaperWidth or PaperHeight (in inches) is larger than the A3 limits (≈11.69 × 16.54 in). Includes sample code that sets custom sizes, runs the check, and saves the workbook.
    public static class PaperSizeHelper
    {
        // A3 size in inches (1 inch = 25.4 mm)
        private const double A3WidthInInches = 297.0 / 25.4;   // ≈ 11.6929
        private const double A3HeightInInches = 420.0 / 25.4; // ≈ 16.5354

        /// <param name="worksheet">The worksheet to examine.</param>
        /// <returns>True if the custom size exceeds A3; otherwise false.</returns>
        public static bool IsCustomPaperSizeExceedsA3(Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            PageSetup pageSetup = worksheet.PageSetup;

            // The paper size must be set to Custom to have a user‑defined size.
            if (pageSetup.PaperSize != PaperSizeType.Custom)
                return false; // Not a custom size, cannot exceed A3 by definition.

            // PaperWidth and PaperHeight are read‑only properties that reflect the current size in inches,
            // taking the page orientation into account.
            double width = pageSetup.PaperWidth;
            double height = pageSetup.PaperHeight;

            // Compare both dimensions with A3 limits.
            // If either dimension is larger, the custom size exceeds A3.
            return width > A3WidthInInches || height > A3HeightInInches;
        }

        // Example usage
        public static void Run()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a custom paper size (e.g., 12 x 18 inches) which exceeds A3.
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;
            sheet.PageSetup.CustomPaperSize(12.0, 18.0);

            bool exceeds = IsCustomPaperSizeExceedsA3(sheet);
            Console.WriteLine($"Custom paper size exceeds A3: {exceeds}");

            // Change to a size within A3 limits (e.g., 10 x 14 inches).
            sheet.PageSetup.CustomPaperSize(10.0, 14.0);
            exceeds = IsCustomPaperSizeExceedsA3(sheet);
            Console.WriteLine($"Custom paper size exceeds A3 after change: {exceeds}");

            // Save the workbook (optional, just to demonstrate lifecycle usage).
            workbook.Save("PaperSizeCheckResult.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                PaperSizeHelper.Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
