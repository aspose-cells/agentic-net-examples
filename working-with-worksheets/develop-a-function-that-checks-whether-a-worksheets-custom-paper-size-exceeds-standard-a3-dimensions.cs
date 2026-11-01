// Title: Determine whether a worksheet’s custom paper size exceeds A3 dimensions using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to detect if worksheet custom paper size is larger than A3 | compare custom PaperWidth and PaperHeight to A3 size with Aspose.Cells PageSetup | Aspose.Cells check if worksheet page setup uses a custom paper size bigger than standard A3 | validate worksheet custom paper dimensions against A3 in a .NET application | detect oversized custom paper size in an Excel file using Aspose.Cells C#
// Tags: custom size exceed check Aspose.Cells | worksheet page setup dimensions validation | A3 dimension comparison C# | PaperWidth PaperHeight analysis Aspose.Cells | detect oversized worksheet print area

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The IsCustomPaperSizeExceedsA3 extension method validates a Worksheet, ensures its PageSetup.PaperSize is set to Custom, then compares the PageSetup.PaperWidth and PaperHeight (points) with the A3 size constants (841.8898 × 1190.5512 points). It returns true when either dimension exceeds the corresponding A3 dimension, allowing developers to programmatically identify oversized custom paper settings.
    public static class WorksheetExtensions
    {
        // A3 dimensions in points (1 point = 1/72 inch)
        // 297 mm = 11.6929 inches * 72 = 841.8898 points
        // 420 mm = 16.5354 inches * 72 = 1190.5512 points
        private const double A3WidthPoints = 841.8898;
        private const double A3HeightPoints = 1190.5512;

        // Checks whether the worksheet's custom paper size exceeds standard A3 dimensions.
        public static bool IsCustomPaperSizeExceedsA3(Worksheet sheet)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));

            var pageSetup = sheet.PageSetup;

            // In Aspose.Cells, custom paper size is represented by the enum value 0.
            // If the paper size is not custom, we cannot compare dimensions, so return false.
            if ((int)pageSetup.PaperSize != 0) // 0 == PaperSize.Custom
                return false;

            // For custom paper size, use PaperWidth and PaperHeight (in points).
            double width = pageSetup.PaperWidth;
            double height = pageSetup.PaperHeight;

            // Determine larger and smaller dimensions to handle portrait/landscape.
            double maxDim = Math.Max(width, height);
            double minDim = Math.Min(width, height);
            double a3Max = Math.Max(A3WidthPoints, A3HeightPoints);
            double a3Min = Math.Min(A3WidthPoints, A3HeightPoints);

            // Return true if either dimension exceeds the corresponding A3 dimension.
            return maxDim > a3Max || minDim > a3Min;
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            // Path to the workbook file (adjust as needed).
            const string workbookPath = "sample.xlsx";

            // Verify that the file exists before attempting to load it.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook.
                var workbook = new Workbook(workbookPath);

                // Use the first worksheet for the demonstration.
                var worksheet = workbook.Worksheets[0];

                // Check if the custom paper size exceeds A3.
                bool exceedsA3 = WorksheetExtensions.IsCustomPaperSizeExceedsA3(worksheet);

                Console.WriteLine($"Custom paper size exceeds A3: {exceedsA3}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
