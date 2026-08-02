// Title: Check if a worksheet’s custom paper size exceeds A3 with Aspose.Cells for .NET
// Description: C# helper that reads a worksheet’s PageSetup, confirms a custom paper size, converts standard A3 (297 × 420 mm) to inches, and returns true when the custom dimensions are larger than A3 regardless of orientation.
// Keywords: Aspose.Cells | C# | custom paper size | A3 dimensions | PageSetup | PaperWidth | PaperHeight | orientation independent | .NET | print layout validation | worksheet size check
// Common Searches: Aspose.Cells check if custom paper size is larger than A3 | C# compare worksheet custom paper dimensions with A3 | How to detect oversized custom paper in Aspose.Cells | PageSetup custom paper size orientation independent Aspose.Cells | Convert A3 mm to inches in Aspose.Cells C#
// Developer Intent: Determine whether a worksheet’s custom paper size is larger than the standard A3 size.
// Use Cases: Validate print layout before exporting to PDF to ensure pages stay within A3 limits. | Enforce size restrictions in reporting tools that only support A3 or smaller sheets. | Log or auto‑adjust worksheets whose custom dimensions exceed A3 for downstream printing pipelines.
// AI Prompts: Generate a C# method that returns true if a worksheet’s custom paper size exceeds A3, handling both portrait and landscape orientations using Aspose.Cells. | Show how to convert A3 dimensions from millimeters to inches and compare them with Worksheet.PageSetup.PaperWidth and PaperHeight. | Explain how to set a custom paper size larger than A3 with Worksheet.PageSetup.CustomPaperSize and then verify it with a helper function.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomPaperCheck
{
    // C# helper that reads a worksheet’s PageSetup, confirms a custom paper size, converts standard A3 (297 × 420 mm) to inches, and returns true when the custom dimensions are larger than A3 regardless of orientation.
    public class PaperSizeHelper
    {
        // Checks if the worksheet uses a custom paper size that is larger than standard A3.
        // Returns true if custom size exceeds A3 dimensions (in any orientation), otherwise false.
        public static bool IsCustomPaperSizeLargerThanA3(Worksheet worksheet)
        {
            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // If the paper size is not set to Custom, there is no custom size to compare.
            if (pageSetup.PaperSize != PaperSizeType.Custom)
                return false;

            // Width and height are returned in inches, taking page orientation into account.
            double customWidth = pageSetup.PaperWidth;
            double customHeight = pageSetup.PaperHeight;

            // A3 size in millimeters: 297 x 420 mm.
            // Convert to inches (1 inch = 25.4 mm).
            const double mmPerInch = 25.4;
            double a3WidthInInches = 297.0 / mmPerInch;   // ≈ 11.69 inches
            double a3HeightInInches = 420.0 / mmPerInch;  // ≈ 16.54 inches

            // Compare dimensions irrespective of orientation.
            double customMax = Math.Max(customWidth, customHeight);
            double customMin = Math.Min(customWidth, customHeight);
            double a3Max = Math.Max(a3WidthInInches, a3HeightInInches);
            double a3Min = Math.Min(a3WidthInInches, a3HeightInInches);

            // If either the larger side or the smaller side exceeds A3, the custom size is larger.
            return customMax > a3Max || customMin > a3Min;
        }
    }

    public class Demo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a custom paper size larger than A3 (e.g., 13 x 18 inches).
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;
            sheet.PageSetup.CustomPaperSize(13.0, 18.0); // width, height in inches

            // Check if the custom size exceeds A3.
            bool exceedsA3 = PaperSizeHelper.IsCustomPaperSizeLargerThanA3(sheet);
            Console.WriteLine("Custom paper size exceeds A3: " + exceedsA3);

            // Save the workbook (optional, demonstrates lifecycle usage).
            workbook.Save("CustomPaperSizeDemo.xlsx");
        }
    }

    // Entry point for testing.
    class Program
    {
        static void Main()
        {
            Demo.Run();
        }
    }
}
