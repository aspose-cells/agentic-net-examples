// Title: Export an Excel workbook with a Gantt chart to PDF while preserving the chart layout using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook containing a Gantt chart and saves it as a PDF with the original chart layout using Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to disable scaling and keep charts unchanged during PDF export.
// Common Searches: Aspose.Cells export Gantt chart to PDF without changing size | C# save Excel workbook as PDF preserving chart layout Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet false for chart export | How to keep Excel chart appearance when converting to PDF with Aspose.Cells | Export Excel Gantt chart to PDF using .NET SDK Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions chart layout preservation | C# export Excel Gantt chart to PDF | Workbook.Save PDF with chart integrity Aspose.Cells | Excel to PDF conversion maintaining charts .NET | Aspose.Cells PDF export settings for Gantt diagrams

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example loads a GanttChart.xlsx workbook, configures PdfSaveOptions with AllColumnsInOnePagePerSheet and OnePagePerSheet set to false to avoid scaling, and saves the workbook as GanttChart.pdf, ensuring the Gantt chart retains its original layout.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "GanttChart.xlsx";
                const string outputFile = "GanttChart.pdf";

                // Ensure the source workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Source file not found: {inputFile}");
                    return;
                }

                // Load the workbook containing the Gantt chart
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options to preserve layout
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    AllColumnsInOnePagePerSheet = false,
                    OnePagePerSheet = false
                };

                // Export the workbook to PDF using the configured options
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"PDF generated successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
