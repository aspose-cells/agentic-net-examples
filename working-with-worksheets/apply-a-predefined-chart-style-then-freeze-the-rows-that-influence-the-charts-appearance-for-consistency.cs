// Title: Apply a predefined chart style and freeze top rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook, add sample sales data, insert a column chart, set its built‑in style to index 9, and freeze the first four rows with Aspose.Cells in C#. | Write C# code that uses Aspose.Cells to apply a predefined chart style to a chart and then lock the top rows so they stay visible while scrolling. | Generate a file named ChartWithStyleAndFrozenRows.xlsx where the chart uses style 9 and the worksheet panes are frozen at row 4 using Aspose.Cells.
// Common Searches: how to set a built‑in chart style in Aspose.Cells C# | freeze panes on specific rows after adding a chart with Aspose.Cells | Aspose.Cells example applying chart style index 9 and freezing top rows | C# Aspose.Cells freeze first four rows of worksheet | apply predefined chart style and freeze header rows in Excel using Aspose.Cells
// Tags: chart style index nine Aspose.Cells | worksheet pane freezing Aspose.Cells | column chart predefined style Aspose.Cells | Excel workbook generation Aspose.Cells C# | header rows visibility Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The program creates a new workbook, populates sample sales data, inserts a column chart, applies chart style index 9, freezes the first four rows of the worksheet, and saves the file as ChartWithStyleAndFrozenRows.xlsx.
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a predefined chart style (style index 9)
            chart.Style = 9;

            // Freeze the first 4 rows (rows 1‑4)
            // The overload requires row, column, totalRows, totalColumns
            sheet.FreezePanes(4, 0, 0, 0);

            // Save the workbook
            workbook.Save("ChartWithStyleAndFrozenRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
