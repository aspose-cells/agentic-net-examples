// Title: Resize chart data label shapes to prevent overflow after applying bold font with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to set chart series data labels to bold and then automatically increase the label shape width and height so the text does not get clipped. | Show how to programmatically adjust the size of Excel chart data label containers after changing the font weight to bold, using the Aspose.Cells .NET API.
// Common Searches: how to increase chart data label size after making text bold with Aspose.Cells C# | Aspose.Cells resize data label shape to avoid overflow in Excel chart | C# Aspose.Cells adjust chart data label dimensions after bold formatting | prevent data label clipping in Excel chart using Aspose.Cells .NET | auto enlarge chart data label container when font is bold Aspose.Cells
// Tags: Aspose.Cells chart data label auto resize | C# bold font overflow fix for Excel chart | adjust chart label shape size with Aspose.Cells | prevent data label clipping in Excel using .NET | Excel chart label dimension scaling Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel workbook, verifies that a chart is present, enables data labels for each series, applies a bold font (and optionally a larger font size) to improve readability, and then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = sheet.Charts[0];

            // Configure data labels for each series
            try
            {
                foreach (Series series in chart.NSeries)
                {
                    // Show data values
                    series.DataLabels.ShowValue = true;

                    // Apply bold font weight
                    series.DataLabels.Font.IsBold = true;

                    // Optional: Adjust font size for readability after bolding
                    series.DataLabels.Font.Size = 12;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error configuring chart series: {ex.Message}");
                return;
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to prepare output directory: {ex.Message}");
                return;
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
