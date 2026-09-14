// Title: Create a timeline for each data row and export the chart to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates over worksheet rows, inserts a Timeline control for each row linked to a pivot table, and saves the resulting column chart as a PDF with Aspose.Cells. | Show how to programmatically add multiple Timeline objects based on row count, then export the associated chart to a PDF file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add a timeline for each data row | Export a column chart to PDF after inserting timelines with Aspose.Cells | Programmatically generate multiple timeline controls linked to a pivot table in Aspose.Cells | Save an Excel workbook containing timelines using Aspose.Cells .NET | How to create timelines from row data in Aspose.Cells
// Tags: add timeline control Aspose.Cells C# | export chart to PDF Aspose.Cells | pivot table timeline creation Aspose.Cells | loop insert timelines Aspose.Cells | save workbook with timelines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Timelines;

// The example creates a new workbook, populates it with date and sales data, builds a pivot table, adds a column chart, iterates over each data row to insert a Timeline control linked to the pivot table, exports the chart as a PDF, and saves the workbook containing all timelines.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Populate sample data (Date and Sales) in the worksheet
            // ------------------------------------------------------------
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales");

            DateTime startDate = new DateTime(2023, 1, 1);
            int dataRows = 10; // number of data rows

            for (int i = 0; i < dataRows; i++)
            {
                // Date column
                cells[i + 1, 0].PutValue(startDate.AddDays(i));
                // Sales column
                cells[i + 1, 1].PutValue(100 + i * 10);
            }

            // ------------------------------------------------------------
            // 2. Create a PivotTable that will serve as the data source for timelines
            // ------------------------------------------------------------
            int pivotIndex = sheet.PivotTables.Add("A1:B11", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh pivot cache and calculate data
            pivot.RefreshData();          // RefreshData is the correct method
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add a chart that visualizes the data (using the original range)
            // ------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Use the original data range for the series values
            chart.NSeries.Add("A1:B11", true);
            // Use the Date column for the category (X) axis
            chart.NSeries.CategoryData = "A2:A11";

            // ------------------------------------------------------------
            // 4. Iterate through each data row and add a Timeline control
            //    Each timeline is placed a few rows below the previous one
            // ------------------------------------------------------------
            int timelineStartRow = 40; // first row where a timeline will be placed
            int timelineColumn = 0;    // column where timelines start

            for (int i = 0; i < dataRows; i++)
            {
                // Add a timeline linked to the same pivot table.
                // Position: (timelineStartRow + i * 5, timelineColumn)
                // The "Date" field is the base field for the timeline.
                sheet.Timelines.Add(pivot, timelineStartRow + i * 5, timelineColumn, "Date");
            }

            // ------------------------------------------------------------
            // 5. Export the chart to a PDF file
            // ------------------------------------------------------------
            string pdfPath = "TimelineChart.pdf";
            chart.ToPdf(pdfPath);

            // ------------------------------------------------------------
            // 6. Save the workbook with all timelines and the chart
            // ------------------------------------------------------------
            string xlsxPath = "TimelineDemo.xlsx";
            workbook.Save(xlsxPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
