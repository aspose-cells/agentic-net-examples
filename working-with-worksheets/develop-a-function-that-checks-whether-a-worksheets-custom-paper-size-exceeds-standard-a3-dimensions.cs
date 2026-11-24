// Title: How to determine whether a worksheet's custom paper size exceeds A3 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# detect custom worksheet paper size larger than A3 | compare worksheet PageSetup PaperWidth and PaperHeight to A3 dimensions in .NET | C# code to check if Excel sheet print size exceeds A3 with Aspose.Cells | determine portrait or landscape overflow of worksheet paper size using Aspose.Cells | validate worksheet print area size against standard A3 in a .NET application
// Tags: worksheet custom paper size validation Aspose.Cells | page setup paper dimensions comparison A3 | detect oversized worksheet print size .NET | Aspose.Cells print area size check | C# worksheet paper size overflow detection

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetPaperSizeApp
{
    // Provides a static helper that reads a worksheet's PageSetup PaperWidth and PaperHeight (in points), compares both portrait and landscape orientations against A3 dimensions, and returns true when the custom size exceeds A3. Includes a sample program that loads or creates a workbook, invokes the helper on the first sheet, and prints the result.
    public static class WorksheetPaperSizeHelper
    {
        // A3 dimensions in points (1 point = 1/72 inch)
        private const double A3WidthPoints = 841.5;
        private const double A3HeightPoints = 1190.5;

        /// <param name="worksheet">The worksheet to examine.</param>
        /// <returns>True if the paper size exceeds A3; otherwise false.</returns>
        public static bool IsPaperSizeExceedingA3(Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            PageSetup pageSetup = worksheet.PageSetup;

            // Retrieve paper dimensions (in points). For standard sizes these values correspond to the preset sizes;
            // for custom sizes they reflect the user‑defined dimensions.
            double width = pageSetup.PaperWidth;
            double height = pageSetup.PaperHeight;

            // Compare both possible orientations (portrait or landscape)
            bool exceedsPortrait = width > A3WidthPoints || height > A3HeightPoints;
            bool exceedsLandscape = width > A3HeightPoints || height > A3WidthPoints;

            return exceedsPortrait || exceedsLandscape;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    // Load existing workbook
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook if the file does not exist
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                Worksheet sheet = workbook.Worksheets[0];

                bool exceeds = WorksheetPaperSizeHelper.IsPaperSizeExceedingA3(sheet);
                Console.WriteLine($"Paper size exceeds A3: {exceeds}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
