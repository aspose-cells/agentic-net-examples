// Title: Create a timeline with dd‑MMM‑yyyy date format, add a column chart, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply dd-MMM-yyyy custom format to timeline control | How to link a timeline to a pivot table and export as PDF with Aspose.Cells | Save Aspose.Cells workbook as PDF/A-1b with chart and timeline in one page | C# example for adding a column chart from pivot data using Aspose.Cells | Export worksheet containing a timeline and chart to PDF using Aspose.Cells .NET | Aspose.Cells timeline control usage with custom date format in C#
// Tags: custom date format timeline Aspose.Cells | PDF/A-1b workbook export Aspose.Cells | pivot table timeline integration C# | column chart from pivot data Aspose.Cells | one page per sheet PDF export Aspose.Cells | timeline control linked to pivot table Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace TimelineChartPdfExport
{
    // The C# program builds a workbook, formats the date column with a custom "dd-MMM-yyyy" pattern, creates a pivot table, adds a timeline linked to that pivot table, generates a column chart from the pivot data, and saves the sheet as a PDF/A‑1b document with the chart and timeline on a single page.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate sample data with dates and sales values
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Ship Date");
                sheet.Cells["B1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["A3"].PutValue(new DateTime(2023, 1, 15));
                sheet.Cells["A4"].PutValue(new DateTime(2023, 2, 1));

                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["B3"].PutValue(2000);
                sheet.Cells["B4"].PutValue(3000);

                // -------------------------------------------------
                // 2. Apply custom date format "dd-MMM-yyyy" to the date column
                // -------------------------------------------------
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "dd-MMM-yyyy";
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                sheet.Cells.CreateRange("A2:A4").ApplyStyle(dateStyle, flag);

                // -------------------------------------------------
                // 3. Create a pivot table based on the data
                // -------------------------------------------------
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "E5", "Pivot1");
                PivotTable pivotTable = sheet.PivotTables[pivotIdx];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Ship Date");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // -------------------------------------------------
                // 4. Add a Timeline control linked to the pivot table
                // -------------------------------------------------
                // The Timeline is placed at cell G5 (row 0, column 0 offset from the sheet origin)
                sheet.Timelines.Add(pivotTable, 0, 0, "Ship Date");

                // -------------------------------------------------
                // 5. Add a chart that visualises the pivot table data
                // -------------------------------------------------
                // Place the chart below the pivot table
                int chartIdx = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
                Chart chart = sheet.Charts[chartIdx];

                // Use the original data range for the chart series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: give the chart a title
                chart.Title.Text = "Sales by Ship Date";

                // -------------------------------------------------
                // 6. Save the workbook as a PDF document
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example: set PDF/A-1b compliance (optional)
                    Compliance = PdfCompliance.PdfA1b,
                    // One page per worksheet ensures the chart and timeline appear together
                    OnePagePerSheet = true
                };

                string outputPath = "TimelineChart.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF successfully created at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
