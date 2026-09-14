// Title: Convert a Waterfall chart Excel workbook to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a .xlsx workbook containing a Waterfall chart, verifies the file exists, and saves it as a PDF using Aspose.Cells. | Demonstrate how to retrieve the first chart from a worksheet before exporting the workbook to PDF with Aspose.Cells in a .NET console application. | Create robust error handling for converting an Excel file with embedded charts to PDF using Aspose.Cells, covering FileNotFoundException and generic exceptions.
// Common Searches: asp.net convert excel file with waterfall chart to pdf using aspose.cells | c# export workbook to pdf preserving charts asp.net | how to save excel workbook as pdf with charts using aspose.cells library | asp.net check if excel file exists before converting to pdf with aspose.cells
// Tags: Aspose.Cells workbook to PDF conversion with charts | C# load Excel workbook and render Waterfall chart to PDF | SaveFormat.Pdf usage in Aspose.Cells | Excel file existence validation in C# before Aspose.Cells conversion | exception handling for Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input Excel file exists, loads the workbook with Aspose.Cells, optionally accesses the first chart on the first worksheet, and then saves the entire workbook—including the Waterfall chart—as a PDF, while providing comprehensive exception handling.
class WaterfallChartToPdf
{
    static void Main()
    {
        try
        {
            // Paths for input workbook and output PDF
            string inputWorkbookPath = "WaterfallChart.xlsx";
            string outputPdfPath = "WaterfallChart.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Input file not found: {inputWorkbookPath}");
                return;
            }

            // Load the workbook containing the Waterfall chart
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Optional: Check if the first worksheet contains any charts
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count > 0)
            {
                // Access the first chart (type inferred with var)
                var waterfallChart = sheet.Charts[0];
                // Additional chart processing can be added here if needed.
            }

            // Convert the entire workbook to PDF (charts are rendered automatically)
            workbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed. PDF saved to: " + outputPdfPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during conversion: " + ex.Message);
        }
    }
}
