// Title: Validate Worksheet Charts Before PDF Export with Aspose.Cells for .NET (C#)
// Description: C# example that creates sample data, adds a column chart, checks Worksheet.Charts.Count, aborts if no chart exists, and exports each chart to a separate PDF using Chart.ToPdf. Includes optional full‑workbook PDF save.
// Keywords: Aspose.Cells C# chart validation | Worksheet.Charts.Count check | Chart.ToPdf example | export chart to PDF Aspose.Cells | Aspose.Cells PDF export .NET | Aspose.Cells USA | Aspose.Cells Europe | Aspose.Cells India | C# Aspose.Cells sample code | PDF generation from Excel charts
// Common Searches: how to verify a worksheet has charts before saving to PDF Aspose.Cells | c# export each chart as separate PDF using Aspose.Cells | prevent empty PDF when no charts in Aspose.Cells worksheet | Aspose.Cells chart to PDF code sample | check worksheet charts count before PDF export
// Developer Intent: Confirm that at least one chart exists on a worksheet before performing any PDF export operations.
// Use Cases: Skip PDF generation when a worksheet contains no charts to avoid empty files. | Export every chart on a worksheet to individual PDF documents. | Save the entire workbook as a PDF after confirming chart presence.
// AI Prompts: Write C# code that checks Worksheet.Charts.Count and calls Chart.ToPdf only when the count is greater than zero. | Show a loop that iterates through all charts in a worksheet and saves each as a separate PDF using Aspose.Cells. | Provide a complete example that adds sample data, creates a chart, validates chart existence, and optionally saves the whole workbook to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartValidation
{
    // C# example that creates sample data, adds a column chart, checks Worksheet.Charts.Count, aborts if no chart exists, and exports each chart to a separate PDF using Chart.ToPdf. Includes optional full‑workbook PDF save.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a potential chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a chart to the worksheet (comment out to test validation when no chart exists)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Validate that the worksheet contains at least one chart before exporting
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("The worksheet does not contain any charts. PDF export aborted.");
                return;
            }

            // Export each chart to a separate PDF file
            for (int i = 0; i < worksheet.Charts.Count; i++)
            {
                Chart c = worksheet.Charts[i];
                string pdfPath = $"Chart_{i + 1}.pdf";
                c.ToPdf(pdfPath);
                Console.WriteLine($"Chart {i + 1} exported to PDF: {pdfPath}");
            }

            // Optionally, export the entire workbook to PDF (charts will be included)
            // workbook.Save("WorkbookWithCharts.pdf", SaveFormat.Pdf);
        }
    }
}
