// Title: Refresh Aspose.Cells Chart After Modifying Source Cells (C#)
// Description: Demonstrates how to update a column chart's source range, call chart.Calculate() to refresh the internal cache, enable PdfSaveOptions.RefreshChartCache for PDF export, and save the workbook as XLSX and PDF so the chart displays the new values.
// Keywords: Aspose.Cells chart refresh | chart.Calculate C# | RefreshChartCache PDF | update chart data programmatically | Aspose.Cells export to PDF | C# Aspose.Cells chart cache | modify chart source cells
// Common Searches: how to refresh a chart in Aspose.Cells after changing cell values | Aspose.Cells chart.Calculate vs RefreshChartCache | export updated chart to PDF with Aspose.Cells | C# update chart source range Aspose.Cells | Aspose.Cells chart not reflecting data changes
// Developer Intent: Ensure the chart reflects the latest cell values before saving the workbook.
// Use Cases: After altering values in the chart's source range, invoke chart.Calculate() to recalculate the chart cache. | Set PdfSaveOptions.RefreshChartCache = true when exporting to PDF to guarantee the PDF uses refreshed chart data. | Save the workbook in both XLSX and PDF formats to verify that the chart displays the updated values in each output.
// AI Prompts: Generate C# code that updates chart source cells, refreshes the chart, and saves the workbook to PDF with the latest data using Aspose.Cells. | Explain the difference between chart.Calculate() and PdfSaveOptions.RefreshChartCache in Aspose.Cells. | Provide a step‑by‑step tutorial for programmatically modifying a chart's data series and ensuring the changes appear in saved Excel and PDF files.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering; // For PdfSaveOptions if needed

namespace AsposeCellsChartRefreshDemo
{
    // Demonstrates how to update a column chart's source range, call chart.Calculate() to refresh the internal cache, enable PdfSaveOptions.RefreshChartCache for PDF export, and save the workbook as XLSX and PDF so the chart displays the new values.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart and set its data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ---- Update source cells ----
            sheet.Cells["B2"].PutValue(15); // Change first value
            sheet.Cells["B3"].PutValue(25); // Change second value
            sheet.Cells["B4"].PutValue(35); // Change third value

            // Refresh the chart so it reflects the updated data
            // For regular charts, calling Calculate updates internal caches
            chart.Calculate();

            // Optionally, when saving to a format that supports chart cache (e.g., PDF),
            // enable RefreshChartCache to ensure the saved output uses the latest data.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.RefreshChartCache = true;
            workbook.Save("ChartRefreshed.pdf", pdfOptions);

            // Also save as an Excel file for verification
            workbook.Save("ChartRefreshed.xlsx");
        }
    }
}
