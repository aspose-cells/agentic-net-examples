using System;
using Aspose.Cells;

namespace AsposeCellsCustomPaperCheck
{
    public class PaperSizeHelper
    {
        // Checks if the worksheet uses a custom paper size that is larger than standard A3 (297mm x 420mm).
        public static bool IsCustomPaperSizeExceedsA3(Worksheet worksheet)
        {
            // Access the page setup of the worksheet.
            PageSetup pageSetup = worksheet.PageSetup;

            // If the paper size is not set to Custom, there is no custom size to evaluate.
            if (pageSetup.PaperSize != PaperSizeType.Custom)
                return false;

            // Retrieve the current custom dimensions (in inches, orientation‑aware).
            double widthInInches = pageSetup.PaperWidth;
            double heightInInches = pageSetup.PaperHeight;

            // A3 dimensions in inches (1 inch = 25.4 mm).
            const double a3WidthInInches = 297.0 / 25.4;   // ≈ 11.6929"
            const double a3HeightInInches = 420.0 / 25.4;  // ≈ 16.5354"

            // Determine if either dimension exceeds the corresponding A3 dimension.
            return widthInInches > a3WidthInInches || heightInInches > a3HeightInInches;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a custom paper size larger than A3 (e.g., 13" x 18").
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;
            sheet.PageSetup.CustomPaperSize(13.0, 18.0); // width, height in inches

            // Use the helper to check if the custom size exceeds A3.
            bool exceedsA3 = PaperSizeHelper.IsCustomPaperSizeExceedsA3(sheet);
            Console.WriteLine("Custom paper size exceeds A3: " + exceedsA3);

            // Save the workbook (optional, demonstrates lifecycle usage).
            workbook.Save("CustomPaperSizeCheck.xlsx");
        }
    }
}