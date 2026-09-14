// Title: Check if a worksheet’s custom paper size is larger than standard A3 using Aspose.Cells for .NET
// AI Prompts: Write a C# function that returns true when the worksheet's PageSetup.PaperWidth or PaperHeight exceeds the A3 size in points, using Aspose.Cells. | Generate code that first verifies the PageSetup.PaperSize is set to Custom, then compares the custom dimensions to the A3 limits and handles any exceptions gracefully. | Provide a complete example that loads an Excel workbook, calls the size‑checking method on the first worksheet, and prints whether the custom paper size is larger than A3.
// Common Searches: Aspose.Cells how to determine if worksheet custom paper size exceeds A3 | C# compare PageSetup.PaperWidth to A3 dimensions in points | detect oversized custom print size in Excel using Aspose.Cells .NET | validate worksheet print setup size against standard A3 with Aspose.Cells | check if Excel worksheet custom paper size is larger than A3 programmatically
// Tags: custom paper size exceeds A3 Aspose.Cells | worksheet page setup size comparison .NET | detect oversized print dimensions C# | validate worksheet print area Aspose.Cells | compare PageSetup.PaperWidth to A3 points

using System;
using System.IO;
using Aspose.Cells;

// A C# utility that inspects a worksheet's PageSetup. If the PaperSize is set to Custom, it compares the PaperWidth and PaperHeight (in points) against the standard A3 dimensions (~842 × 1190 points) and returns a boolean, with safe exception handling and a usage example.
public class PaperSizeChecker
{
    // A3 dimensions in points (1 point = 1/72 inch)
    private const double A3WidthPoints = 297.0 * 72.0 / 25.4;   // ≈ 842.0 points
    private const double A3HeightPoints = 420.0 * 72.0 / 25.4; // ≈ 1190.5 points

    /// <param name="worksheet">The worksheet to examine.</param>
    /// <returns>True if custom size exceeds A3; otherwise false.</returns>
    public static bool IsCustomPaperSizeExceedsA3(Worksheet worksheet)
    {
        try
        {
            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // If the paper size is not set to Custom, there is no custom size to compare
            // Use string comparison to avoid direct enum reference (compatible with all versions)
            if (!pageSetup.PaperSize.ToString().Equals("Custom", StringComparison.OrdinalIgnoreCase))
                return false;

            // Retrieve custom width and height (in points)
            double customWidth = pageSetup.PaperWidth;
            double customHeight = pageSetup.PaperHeight;

            // Compare with A3 dimensions
            return customWidth > A3WidthPoints || customHeight > A3HeightPoints;
        }
        catch (Exception ex)
        {
            // Log and treat any error as non‑exceeding to keep the method safe
            Console.WriteLine($"Error while checking paper size: {ex.Message}");
            return false;
        }
    }

    // Example usage
    public static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            bool exceedsA3 = IsCustomPaperSizeExceedsA3(sheet);
            Console.WriteLine($"Custom paper size exceeds A3: {exceedsA3}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
