// Title: Export Gantt Chart from Excel to PDF with Layout Preservation using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, verifies that the first worksheet contains a chart, and uses Aspose.Cells' Chart.ToPdf method to export the Gantt chart to a PDF. The export applies a standard Letter page size and centers the chart both horizontally and vertically, with graceful handling of missing files or charts.
// Keywords: Aspose.Cells | C# | .NET | export chart to PDF | Gantt chart PDF | Chart.ToPdf | preserve chart layout | page size alignment | Excel to PDF conversion | Aspose.Cells chart export
// Common Searches: Aspose.Cells export Gantt chart to PDF C# | Chart.ToPdf page size alignment example | How to keep chart layout when converting Excel to PDF | Export specific Excel chart as PDF using Aspose.Cells | C# code to export Excel chart with custom dimensions
// Developer Intent: Export a Gantt chart from an Excel workbook to a PDF while retaining its original layout and alignment.
// Use Cases: Generate a PDF report of a project schedule by exporting the Gantt chart from a template workbook. | Automate batch processing of multiple workbooks, converting each Gantt chart to a separate PDF with consistent page settings. | Create a custom PDF export routine that adjusts page width, height, and alignment for different chart types to ensure visual fidelity.
// AI Prompts: Write C# code that uses Aspose.Cells to export a selected chart from an Excel file to a PDF with custom page dimensions and centered alignment. | Provide error‑handling best practices for chart‑to‑PDF conversion with Aspose.Cells, including file existence and chart count checks. | Show how to loop through all worksheets in a workbook and save each chart as an individual PDF while preserving layout.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartExport
{
    // Loads an Excel workbook, verifies that the first worksheet contains a chart, and uses Aspose.Cells' Chart.ToPdf method to export the Gantt chart to a PDF. The export applies a standard Letter page size and centers the chart both horizontally and vertically, with graceful handling of missing files or charts.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "GanttChart.xlsx";
                const string outputPath = "GanttChart.pdf";

                // Verify that the source workbook exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook that already contains the Gantt chart.
                Workbook workbook = new Workbook(inputPath);

                // Assume the Gantt chart is the first chart on the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the workbook.");
                    return;
                }

                Chart ganttChart = worksheet.Charts[0];

                // Export the chart to PDF while preserving its layout.
                // Page size set to standard Letter (8.5 x 11 inches) and centered on the page.
                ganttChart.ToPdf(
                    outputPath,
                    8.5f,                     // desired page width in inches
                    11f,                      // desired page height in inches
                    PageLayoutAlignmentType.Center, // horizontal alignment
                    PageLayoutAlignmentType.Center  // vertical alignment
                );

                Console.WriteLine($"Gantt chart exported to PDF successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
