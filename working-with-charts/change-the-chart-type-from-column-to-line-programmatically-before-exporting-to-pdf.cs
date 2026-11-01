// Title: Convert column and stacked column charts to line charts programmatically before saving an Excel workbook as PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX file, changes every Column or StackedColumn chart to a Line chart, and saves the workbook as a PDF using Aspose.Cells. | Generate a script that iterates through all worksheets in a workbook, updates chart.Type from Column/ColumnStacked to Line, then exports the result to PDF. | Provide a .NET example that replaces column‑type charts with line charts in an existing workbook and creates a PDF output.
// Common Searches: how to change column chart to line chart with Aspose.Cells C# before PDF export | Aspose.Cells .NET replace stacked column chart with line chart programmatically | convert Excel chart types to line and generate PDF using Aspose.Cells | C# iterate workbook charts and modify type prior to saving as PDF
// Tags: chart type conversion using Aspose.Cells | programmatic chart replacement before PDF generation | Aspose.Cells workbook chart manipulation example | C# export Excel to PDF after chart changes | update Excel chart objects with .NET API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code loads an existing XLSX workbook, loops through each worksheet and its charts, converts any Column or Stacked Column chart to a Line chart, and then saves the modified workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Convert column charts (including stacked) to line charts
                    if (chart.Type == ChartType.Column ||
                        chart.Type == ChartType.ColumnStacked)
                    {
                        chart.Type = ChartType.Line;
                    }
                }
            }

            // Export the modified workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
