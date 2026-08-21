// Title: Add Multiple Series to a Column Chart and Freeze Data Rows with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, populates columns A‑D with category labels and three numeric series, adds a column chart that references each series, assigns the category (X‑axis) range, freezes rows 1‑6 using FreezePanes, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | column chart | multiple series | add series to chart | freeze rows | freeze panes | NSeries | category axis | Excel automation | XLSX generation
// Common Searches: Aspose.Cells add several series to a column chart | How to freeze specific rows in an Aspose.Cells worksheet | Set category axis data for a chart using Aspose.Cells | Create chart with data from multiple columns in C# | Freeze panes programmatically with Aspose.Cells
// Developer Intent: Programmatically build a column chart with three data series and keep the source rows visible by freezing them.
// Use Cases: Sales dashboard where each product line appears as a separate series and the underlying data rows stay in view while scrolling. | Monthly performance report that charts key metrics and freezes the data rows for quick reference during analysis. | Presentation workbook that generates charts on the fly while keeping the source data rows locked for easy editing.
// AI Prompts: Generate C# code to add five series to a line chart and freeze the first ten rows with Aspose.Cells. | Show how to set category data from a named range and freeze both rows and columns in an Aspose.Cells worksheet. | Explain how to add chart series dynamically based on the number of data columns and keep the data rows frozen.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a new workbook, populates columns A‑D with category labels and three numeric series, adds a column chart that references each series, assigns the category (X‑axis) range, freezes rows 1‑6 using FreezePanes, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data:
            // Column A – categories, Columns B‑D – three separate series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            string[] categories = { "A", "B", "C", "D", "E" };
            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // data starts at row 2
                sheet.Cells[$"A{row}"].PutValue(categories[i]);
                sheet.Cells[$"B{row}"].PutValue((i + 1) * 10); // Series1 values
                sheet.Cells[$"C{row}"].PutValue((i + 1) * 15); // Series2 values
                sheet.Cells[$"D{row}"].PutValue((i + 1) * 20); // Series3 values
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add three series to the chart (by column, vertical data)
            chart.NSeries.Add("=Sheet1!$B$2:$B$6", true); // Series1
            chart.NSeries.Add("=Sheet1!$C$2:$C$6", true); // Series2
            chart.NSeries.Add("=Sheet1!$D$2:$D$6", true); // Series3

            // Set the category (X‑axis) data for all series
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

            // Freeze rows 1‑6 (first unfrozen row is 7). Column index 0 means no column freeze.
            // totalRows and totalColumns set to 0 to indicate default pane size.
            sheet.FreezePanes(7, 0, 0, 0);

            // Save the workbook
            string outputPath = "MultipleSeries_FrozenRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
