// Title: Export a Progress Bar chart to a JPEG file with a white background using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export chart as JPEG with white background for email | How to save a progress bar chart from an Excel workbook as a JPEG image using Aspose.Cells | Aspose.Cells chart.ToImage JPEG white background example in .NET | Generate JPEG image of a worksheet chart for email attachment with Aspose.Cells | Export first worksheet chart to JPEG file and set background color using Aspose.Cells C#
// Tags: Aspose.Cells chart to JPEG conversion | progress bar chart image export .NET | chart image white background Aspose.Cells | C# export Excel chart as JPEG file | email attachment chart image generation

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook, selects the first chart on the first worksheet, converts the chart to a JPEG image with a white background using Chart.ToImage, and writes the image to a specified file path, handling missing files, conversion errors, and write failures.
class ExportProgressBarChart
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = @"C:\Charts\ProgressBarChart.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found at path: {workbookPath}");
                return;
            }

            // Load the workbook that contains the progress bar chart
            using (Workbook workbook = new Workbook(workbookPath))
            {
                // Assume the progress bar chart is the first chart on the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                Chart chart = sheet.Charts[0];

                // Export the chart as a JPEG image
                using (MemoryStream ms = new MemoryStream())
                {
                    try
                    {
                        // Save the chart image to the memory stream in JPEG format
                        chart.ToImage(ms, ImageFormat.Jpeg);
                    }
                    catch (Exception imgEx)
                    {
                        Console.WriteLine($"Failed to convert chart to image: {imgEx.Message}");
                        return;
                    }

                    // Write the JPEG bytes to a file
                    string outputPath = @"C:\Charts\ProgressBarChart.jpg";

                    // Ensure the output directory exists
                    string? outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    try
                    {
                        File.WriteAllBytes(outputPath, ms.ToArray());
                        Console.WriteLine($"Chart exported successfully to: {outputPath}");
                    }
                    catch (Exception writeEx)
                    {
                        Console.WriteLine($"Failed to write image file: {writeEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
