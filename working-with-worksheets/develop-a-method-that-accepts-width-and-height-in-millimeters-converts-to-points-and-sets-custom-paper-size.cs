// Title: Set Custom Worksheet Paper Size in Points from Millimeters with Aspose.Cells for .NET
// Description: C# method that receives width and height in millimeters, converts them to inches (1 in = 25.4 mm) and points (1 in = 72 pt), applies the dimensions to the first worksheet via PageSetup.CustomPaperSize, ensures the target folder exists, and saves the workbook.
// Keywords: Aspose.Cells custom paper size | C# convert mm to points | Excel worksheet page setup inches | Aspose.Cells set page size programmatically | metric to point conversion .NET | custom worksheet dimensions | Aspose.Cells save workbook path
// Common Searches: Aspose.Cells set custom paper size C# | convert millimeters to points for Excel page setup | how to define custom worksheet size in Aspose.Cells | C# create workbook with A4 dimensions using Aspose.Cells | set page size in points instead of inches Aspose.Cells
// Developer Intent: Create a reusable C# routine that takes metric dimensions, converts them to the units required by Aspose.Cells, and assigns a custom paper size to a worksheet.
// Use Cases: Printing labels or forms where the paper size is specified in millimeters. | Generating regional reports (A4, A5, Legal) directly from metric specifications. | Preparing Excel files for PDF conversion with exact point measurements to match print layouts.
// AI Prompts: Generate a C# function using Aspose.Cells that accepts width and height in millimeters, converts them to inches and points, and sets the worksheet's custom paper size. | Add comprehensive error handling to the custom paper size method, including directory creation and detailed logging of conversion failures. | Show how to retrieve the calculated point values after setting the custom size and use them to adjust worksheet margins programmatically.

using System;
using System.IO;
using Aspose.Cells;

// C# method that receives width and height in millimeters, converts them to inches (1 in = 25.4 mm) and points (1 in = 72 pt), applies the dimensions to the first worksheet via PageSetup.CustomPaperSize, ensures the target folder exists, and saves the workbook.
public static class PaperSizeHelper
{
    /// <param name="widthMm">Paper width in millimeters.</param>
    /// <param name="heightMm">Paper height in millimeters.</param>
    /// <param name="outputPath">Full path where the workbook will be saved.</param>
    public static void SetCustomPaperSizeInPoints(double widthMm, double heightMm, string outputPath)
    {
        try
        {
            // Convert millimeters to inches (1 inch = 25.4 mm)
            double widthInches = widthMm / 25.4;
            double heightInches = heightMm / 25.4;

            // For informational purposes: convert inches to points (1 inch = 72 points)
            double widthPoints = widthInches * 72.0;
            double heightPoints = heightInches * 72.0;

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the custom paper size using inches (required by the API)
            worksheet.PageSetup.CustomPaperSize(widthInches, heightInches);

            // Optionally, you could store the point values somewhere or use them for other calculations
            // Console.WriteLine($"Custom size in points: {widthPoints}pt x {heightPoints}pt");

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error setting custom paper size: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Example usage: A4 size (210mm x 297mm) saved to "CustomPaperSize.xlsx"
            double widthMm = 210;
            double heightMm = 297;
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomPaperSize.xlsx");

            PaperSizeHelper.SetCustomPaperSizeInPoints(widthMm, heightMm, outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
