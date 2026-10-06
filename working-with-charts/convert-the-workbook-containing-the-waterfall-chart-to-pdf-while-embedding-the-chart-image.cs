// Title: Convert a Waterfall chart from an Excel workbook to a PDF by embedding it as an image using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file, finds the first Waterfall chart, uses Chart.ToImage to export it to PNG, inserts the PNG into a new worksheet, and saves the result as a PDF with Aspose.Cells. | Show how to detect the absence of a Waterfall chart in a workbook and return a friendly message while performing the conversion with Aspose.Cells. | Demonstrate proper cleanup of a temporary PNG file after converting an Excel chart to PDF in a C# application.
// Common Searches: Aspose.Cells C# convert specific Waterfall chart to PDF | How to export only a chart from Excel to PDF using Aspose.Cells | Render Excel Waterfall chart as image and embed in PDF programmatically | C# sample for saving Excel chart as PNG then creating PDF with Aspose | Temporary file handling when converting Excel chart to PDF Aspose.Cells
// Tags: Chart.ToImage method Aspose.Cells | Save workbook as PDF Aspose.Cells | Insert picture into worksheet Aspose.Cells | Waterfall chart detection Aspose.Cells | Temporary PNG cleanup C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads WaterfallChart.xlsx, locates the first Waterfall chart, renders it to a temporary PNG file, inserts the image into a new workbook, saves that workbook as WaterfallChart.pdf, and finally deletes the temporary image.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "WaterfallChart.xlsx";
            const string outputPdf = "WaterfallChart.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook that contains the Waterfall chart
            Workbook workbook = new Workbook(inputFile);

            // Assume the chart is on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the first Waterfall chart in the worksheet
            Chart waterfallChart = null;
            foreach (Chart chart in worksheet.Charts)
            {
                if (chart.Type == ChartType.Waterfall)
                {
                    waterfallChart = chart;
                    break;
                }
            }

            if (waterfallChart == null)
            {
                Console.WriteLine("No Waterfall chart found in the workbook.");
                return;
            }

            // Create a temporary PNG file for the chart image
            string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");

            try
            {
                // Render the chart directly to the temporary PNG file
                waterfallChart.ToImage(tempImagePath);

                // Create a temporary workbook to hold the chart image
                Workbook imgWorkbook = new Workbook();
                Worksheet imgSheet = imgWorkbook.Worksheets[0];

                // Insert the chart image into the worksheet
                imgSheet.Pictures.Add(0, 0, tempImagePath);

                // Save the temporary workbook as PDF (contains only the chart image)
                imgWorkbook.Save(outputPdf, SaveFormat.Pdf);
            }
            finally
            {
                // Clean up the temporary image file
                if (File.Exists(tempImagePath))
                {
                    try { File.Delete(tempImagePath); } catch { /* ignore cleanup errors */ }
                }
            }

            Console.WriteLine($"PDF file '{outputPdf}' created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
