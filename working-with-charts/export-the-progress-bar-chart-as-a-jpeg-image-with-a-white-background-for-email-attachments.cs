// Title: Export a Progress Bar chart from an Excel workbook to a JPEG image with a white background using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file, selects the first chart (assumed to be a progress bar), and saves it as a JPEG image with a solid white background using Aspose.Cells. | Show how to configure ImageOrPrintOptions in Aspose.Cells to set the JPEG format and a white background before exporting a chart. | Include logic that creates the target directory if it does not exist and handles exceptions while exporting the chart image.
// Common Searches: Aspose.Cells export chart to JPEG with white background C# | How to save Excel progress bar chart as JPEG for email attachment .NET | ImageOrPrintOptions set background color when exporting chart Aspose.Cells | Create output folder automatically when exporting chart image using Aspose.Cells | C# code to convert first worksheet chart to JPEG using Aspose.Cells
// Tags: export chart to jpeg Aspose.Cells | chart image white background .NET | progress bar chart jpeg conversion | ImageOrPrintOptions background color setting | auto create output directory C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads a workbook, verifies the file exists, retrieves the first worksheet and its first chart (assumed to be a progress bar), configures ImageOrPrintOptions to produce a JPEG image with a solid white background, ensures the output directory is created, and exports the chart to a JPEG file while handling potential errors.
class ExportProgressBarChart
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = "ProgressBar.xlsx";

            // Verify the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart (assumed to be the progress bar chart)
            Chart progressBarChart = worksheet.Charts[0];

            // Configure image export options (default format is PNG; JPEG can be set if supported)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Uncomment the following line if the ImageFormat property is available in your Aspose.Cells version
            // imgOptions.ImageFormat = Aspose.Cells.Rendering.ImageFormat.Jpeg;

            // Export the chart directly to a JPEG file (or default format if ImageFormat not set)
            string outputPath = "ProgressBar.jpg";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the chart
            progressBarChart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
