// Title: Set chart data source range and freeze source rows in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart series range from worksheet cells | How to freeze rows that contain chart source data using Aspose.Cells | C# example for linking column chart to a table and freezing header rows | Freeze panes after creating a chart with Aspose.Cells .NET | Set category labels and values for a column chart in Aspose.Cells
// Tags: chart series range definition Aspose.Cells | freeze panes rows worksheet Aspose.Cells | column chart data binding Aspose.Cells .NET | set category axis labels Aspose.Cells | save workbook as XLSX Aspose.Cells | populate worksheet table for chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills a table with category and value data, adds a column chart whose series and category axes reference that table, freezes the rows containing the source data to keep them visible while scrolling, and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // -------------------------------------------------
            // 1. Populate source data that will feed the chart
            // -------------------------------------------------
            // Header row
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");

            // Sample data rows (rows 2‑6)
            for (int row = 2; row <= 6; row++)
            {
                worksheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                worksheet.Cells[$"B{row}"].PutValue(row * 10); // arbitrary numeric value
            }

            // -------------------------------------------------
            // 2. Add a chart and set its data source range
            // -------------------------------------------------
            // Add a column chart positioned below the data table
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 8, 0, 25, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the series (Y‑values) range – column B rows 2‑6
            chart.NSeries.Add("=DataSheet!$B$2:$B$6", true);

            // Set the category (X‑axis) labels – column A rows 2‑6
            chart.NSeries.CategoryData = "=DataSheet!$A$2:$A$6";

            // Optional: give the chart a title
            chart.Title.Text = "Sample Column Chart";

            // -------------------------------------------------
            // 3. Freeze the rows that contain the source data
            // -------------------------------------------------
            // Freeze the first 6 rows (rows 1‑6) so they stay visible while scrolling
            // FreezePanes(string cellName, int totalRows, int totalColumns)
            // CellName is the top‑left cell of the unfrozen pane; using "A7" freezes rows 1‑6.
            worksheet.FreezePanes("A7", 6, 0);

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            string outputPath = "ChartWithFrozenSource.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
