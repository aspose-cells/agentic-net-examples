// Title: Add exception handling for invalid custom paper dimensions when setting a custom page size with Aspose.Cells for .NET
// AI Prompts: Validate that the width and height values are positive before calling Sheet.PageSetup.CustomPaperSize, and throw an ArgumentException with a clear message if they are not. | Enclose the custom paper size configuration in a try‑catch block that captures any exception, logs a descriptive error to the console or a logger, and continues or aborts gracefully.
// Common Searches: c# Aspose.Cells validate custom paper size before applying | Aspose.Cells custom page setup exception handling example | how to log error when setting custom paper dimensions in Aspose.Cells | prevent negative dimensions for custom paper size Aspose.Cells .NET | save workbook after catching custom paper size errors Aspose.Cells
// Tags: custom paper size validation Aspose.Cells | Aspose.Cells page setup exception handling | Aspose.Cells workbook save error logging | Aspose.Cells custom dimensions .NET | Aspose.Cells PageSetup.CustomPaperSize error handling

using System;
using System.IO;
using System.Drawing; // For SizeF if needed
using Aspose.Cells;

// The example creates a workbook, checks that custom paper width and height are positive, converts inches to points, applies a custom paper size to the first worksheet, ensures the output directory exists, saves the file, and logs any exceptions that occur during the custom size configuration.
class CustomPaperExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Desired dimensions in inches
            double widthInches = 8.5;
            double heightInches = 11.0;

            // Validate dimensions
            if (widthInches <= 0 || heightInches <= 0)
                throw new ArgumentException("Paper dimensions must be positive numbers.");

            // Convert inches to points (1 inch = 72 points)
            float widthPoints = (float)(widthInches * 72);
            float heightPoints = (float)(heightInches * 72);

            // Apply custom paper size
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;

            // Set custom paper size using the appropriate API (method overload)
            sheet.PageSetup.CustomPaperSize(widthPoints, heightPoints);

            // Define output path and ensure directory exists
            string outputPath = "CustomPaperOutput.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any errors that occur
            Console.WriteLine($"Error setting custom paper dimensions: {ex.Message}");
        }
    }
}
