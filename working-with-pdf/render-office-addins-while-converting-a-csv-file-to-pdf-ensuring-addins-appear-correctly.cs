// Title: Render Office Add‑In (Chart) When Converting CSV to PDF with Aspose.Cells for .NET
// Description: This example creates a temporary CSV file, loads it into an Aspose.Cells Workbook, adds a column chart to simulate an Office Add‑In, enables RefreshChartCache in PdfSaveOptions, and saves the workbook as a PDF so the chart appears correctly. The temporary CSV is then removed.
// Keywords: Aspose.Cells CSV to PDF | render chart in PDF | RefreshChartCache | Office Add‑In export | C# Aspose.Cells PDF conversion | add chart before PDF save | .NET workbook to PDF
// Common Searches: Aspose.Cells include chart when converting CSV to PDF | PdfSaveOptions RefreshChartCache example | convert CSV file to PDF with chart using C# | how to render Office Add‑In in PDF export Aspose.Cells | add chart to workbook loaded from CSV before PDF export
// Developer Intent: Add a chart that represents an Office Add‑In to a workbook loaded from a CSV file and export the workbook to PDF with the chart rendered accurately.
// Use Cases: Generate PDF reports from CSV data that contain visual charts for business dashboards. | Create printable invoices where sales figures imported from CSV are displayed as a column chart. | Automate conversion of CSV log files into PDF summaries that embed charts for quick insight.
// AI Prompts: Provide C# code that loads a CSV into Aspose.Cells, adds a column chart, sets PdfSaveOptions.RefreshChartCache, and saves as PDF. | Explain the impact of RefreshChartCache on chart rendering during PDF export with Aspose.Cells. | Step‑by‑step guide to ensure charts added after loading a CSV appear in the final PDF using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;
using Aspose.Cells.Charts;

namespace AsposeCellsAddInPdfConversion
{
    // This example creates a temporary CSV file, loads it into an Aspose.Cells Workbook, adds a column chart to simulate an Office Add‑In, enables RefreshChartCache in PdfSaveOptions, and saves the workbook as a PDF so the chart appears correctly. The temporary CSV is then removed.
    class Program
    {
        static void Main()
        {
            // Paths for temporary CSV and final PDF
            string csvPath = "sample_data.csv";
            string pdfPath = "output_with_addins.pdf";

            // ------------------------------------------------------------
            // 1. Create a sample CSV file (this is the source file)
            // ------------------------------------------------------------
            File.WriteAllText(csvPath,
                "Category,Value\n" +
                "Fruits,50\n" +
                "Vegetables,30\n" +
                "Grains,20");

            // ------------------------------------------------------------
            // 2. Load the CSV into a Workbook (using LoadOptions)
            // ------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions); // create + load

            // ------------------------------------------------------------
            // 3. Add a chart – this represents an Office Add‑In that must
            //    appear in the final PDF. The chart is added to the first
            //    worksheet after the CSV data has been imported.
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            // Set data range for the chart (B2:B4 contains the values)
            chart.NSeries.Add("B2:B4", true);
            // Set category labels (A2:A4 contains the categories)
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Category Chart";

            // ------------------------------------------------------------
            // 4. Prepare PDF save options.
            //    RefreshChartCache ensures that the chart is rendered correctly.
            // ------------------------------------------------------------
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                RefreshChartCache = true
            };

            // ------------------------------------------------------------
            // 5. Save the workbook as PDF (single create‑save operation)
            // ------------------------------------------------------------
            workbook.Save(pdfPath, pdfSaveOptions);

            // ------------------------------------------------------------
            // 6. Clean up temporary CSV file (optional)
            // ------------------------------------------------------------
            if (File.Exists(csvPath))
                File.Delete(csvPath);

            Console.WriteLine($"CSV converted to PDF with chart add‑in: {pdfPath}");
        }
    }
}
