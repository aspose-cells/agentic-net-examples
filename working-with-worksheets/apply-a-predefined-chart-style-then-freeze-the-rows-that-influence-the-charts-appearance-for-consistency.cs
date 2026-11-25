// Title: Apply a predefined chart style and freeze top rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply chart style index 1 to a column chart | Freeze first two rows in an Excel worksheet with Aspose.Cells .NET | Set predefined chart style and then freeze panes after creating a chart using Aspose.Cells | C# add column chart with data source and freeze header rows in Excel via Aspose.Cells | How to use FreezePanes method after chart creation in Aspose.Cells for .NET
// Tags: apply predefined chart style Aspose.Cells | freeze top rows worksheet Aspose.Cells | add column chart with data source Aspose.Cells | set chart title Aspose.Cells C# | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartStyleAndFreeze
{
    // The example creates a workbook, fills it with month and sales data, adds a column chart, applies predefined chart style 1, freezes the first two rows to keep headers visible, and saves the file as 'ChartWithStyleAndFrozenRows.xlsx' using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // -----------------------------------------------------------------
                // 2. Populate sample data that will be used for the chart.
                //    Column A – Category (Month), Column B – Values.
                // -----------------------------------------------------------------
                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] values = { 1200, 1500, 1800, 1300, 1700 };

                // Header row
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");

                // Data rows
                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A (index 0)
                    sheet.Cells[i + 2, 1].PutValue(values[i]); // Column B (index 1)
                }

                // -----------------------------------------------------------------
                // 3. Add a chart to the worksheet.
                // -----------------------------------------------------------------
                // The chart will be placed with its upper‑left corner at cell D2
                // and will occupy a width of 15 columns and a height of 15 rows.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 1, 3, 16, 18);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Monthly Sales";

                // Set the data source for the chart.
                // Category (X‑axis) – A2:A6, Values (Y‑axis) – B2:B6
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // -----------------------------------------------------------------
                // 4. Apply a predefined chart style.
                // -----------------------------------------------------------------
                // Aspose.Cells uses an integer index for chart styles (1‑48).
                // Here we apply style 1.
                chart.Style = 1;

                // -----------------------------------------------------------------
                // 5. Freeze the rows that influence the chart’s appearance.
                //    Freeze the first two rows (header + data rows).
                // -----------------------------------------------------------------
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Row/column specify the cell where the split occurs (zero‑based).
                // We split at row index 2 (third row) to freeze rows 0‑1.
                sheet.FreezePanes(2, 0, 2, 0); // Freeze top 2 rows, no columns.

                // -----------------------------------------------------------------
                // 6. Save the workbook.
                // -----------------------------------------------------------------
                string outputPath = "ChartWithStyleAndFrozenRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Report any unexpected errors.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
