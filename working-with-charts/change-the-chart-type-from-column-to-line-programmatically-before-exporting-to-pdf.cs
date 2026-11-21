// Title: Change column charts to line charts programmatically and export the workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change chart type from column to line before saving as PDF | How to replace column chart with line chart in an Excel file using Aspose.Cells .NET | Programmatically modify Excel chart type and export to PDF with Aspose.Cells | Iterate worksheets and convert column charts to line charts in C# | Save workbook as PDF after updating chart types using Aspose.Cells
// Tags: Aspose.Cells chart type conversion column to line | C# modify Excel chart before PDF export | Aspose.Cells change chart type programmatically | export workbook to PDF Aspose.Cells | iterate worksheets charts Aspose.Cells C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an existing Excel workbook, iterates through each worksheet's charts, changes any Column chart to a Line chart, and then saves the modified workbook directly as a PDF.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Change Column charts to Line charts
                    if (chart.Type == ChartType.Column)
                    {
                        chart.Type = ChartType.Line;
                    }
                }
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
