using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace GanttChartPdfExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a Gantt‑like chart
                // Columns: Task, Start Date, End Date
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("End");

                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 15));

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 16));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 2, 28));

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 3, 15));

                // Add a stacked bar chart (used to mimic a Gantt chart)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 8);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                // Series: Start and End dates
                ganttChart.NSeries.Add("B2:C4", true);
                // Category (Task) labels
                ganttChart.NSeries.CategoryData = "A2:A4";

                // Optional: set chart title
                ganttChart.Title.Text = "Project Schedule";

                // Prepare PDF save options to preserve chart layout and document structure
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true,   // preserve layout information
                    RefreshChartCache = true          // ensure chart is rendered with latest data
                };

                // Define output file path
                string outputPath = "GanttChart.pdf";

                // Save the workbook (including the chart) as a PDF file
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook with chart exported to PDF successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}