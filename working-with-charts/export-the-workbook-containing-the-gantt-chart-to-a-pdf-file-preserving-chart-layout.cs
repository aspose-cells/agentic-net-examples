// Title: Export an Excel workbook that contains a Gantt chart to PDF while preserving the chart layout using Aspose.Cells for .NET
// Common Searches: c# aspnet cells export excel file with gantt chart to pdf preserving layout | how to keep chart formatting when saving workbook as pdf using Aspose.Cells | pdfsaveoptions settings to avoid scaling charts in Aspose.Cells .NET | export gantt chart from excel to pdf without losing column widths asp.net cells | save workbook as pdf with original chart size aspose.cells c#
// Tags: Aspose.Cells PdfSaveOptions chart preservation | C# export Excel Gantt chart to PDF | PDF export without chart scaling Aspose.Cells | retain column widths in PDF save Aspose.Cells | render Excel charts in PDF .NET

using System;
using System.IO;
using Aspose.Cells;

namespace GanttChartExport
{
    // The example loads an existing Excel workbook that contains a Gantt chart, configures PdfSaveOptions with AllColumnsInOnePagePerSheet and OnePagePerSheet set to false to keep original column widths, row heights, and chart rendering, and then saves the workbook as a PDF file while preserving the chart layout.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "GanttChartWorkbook.xlsx";
            string outputPath = "GanttChartExport.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found - {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook that contains the Gantt chart
                var workbook = new Workbook(inputPath);

                // Configure PDF save options to preserve the original layout (including charts)
                var pdfSaveOptions = new PdfSaveOptions
                {
                    // Keep the original column widths and row heights; do not force everything onto one page
                    AllColumnsInOnePagePerSheet = false,
                    // Ensure that charts are rendered as they appear in the worksheet
                    OnePagePerSheet = false
                    // Note: ImageQuality property is not available in current Aspose.Cells versions
                };

                // Export the workbook to PDF while preserving the Gantt chart layout
                workbook.Save(outputPath, pdfSaveOptions);
                Console.WriteLine($"Workbook successfully exported to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }
    }
}
