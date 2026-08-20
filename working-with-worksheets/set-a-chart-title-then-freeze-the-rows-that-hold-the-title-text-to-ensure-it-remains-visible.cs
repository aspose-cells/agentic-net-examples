// Title: Aspose.Cells for .NET: Set Chart Title from a Cell and Freeze the Title Row (C#)
// Description: Creates a workbook, writes a title in A1, adds sample data, inserts a column chart, uses the cell value as the chart title, makes the title visible, and freezes row 1 with PageSetup.PrintTitleRows before saving as ChartWithFrozenTitleRow.xlsx.
// Keywords: Aspose.Cells chart title from cell | freeze row Aspose.Cells C# | PageSetup.PrintTitleRows example | Excel header freeze Aspose.Cells | C# Aspose.Cells chart title visibility
// Common Searches: Aspose.Cells set chart title from worksheet cell | freeze first row in Aspose.Cells workbook | C# Aspose.Cells PrintTitleRows usage | how to keep chart title row visible in generated Excel | Aspose.Cells chart title and row freeze
// Developer Intent: Assign a worksheet cell value as a chart title and keep the title row fixed while scrolling or printing.
// Use Cases: Generate a sales‑summary column chart where the title is driven by cell A1 and the header row stays visible on screen and printed pages. | Automate Excel reports that need a dynamic chart title and a locked title row for consistent branding. | Create an export routine that adds a chart, pulls its title from a cell, and applies PrintTitleRows to repeat the title on each printed sheet.
// AI Prompts: Provide C# Aspose.Cells code that reads a cell for the chart title, adds a column chart, and freezes the title row using PrintTitleRows. | Show how to set PageSetup.PrintTitleRows to "$1:$1" after assigning a chart title from a worksheet cell in Aspose.Cells. | Write an Aspose.Cells example that creates a chart, uses a merged cell as the title, and keeps the title row visible when scrolling.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleFreeze
{
    // Creates a workbook, writes a title in A1, adds sample data, inserts a column chart, uses the cell value as the chart title, makes the title visible, and freezes row 1 with PageSetup.PrintTitleRows before saving as ChartWithFrozenTitleRow.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put the chart title text into a cell (row 1) – this row will be frozen
            sheet.Cells["A1"].PutValue("Sales Overview");

            // Add sample data for the chart
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");
            sheet.Cells["B2"].PutValue(15000);
            sheet.Cells["B3"].PutValue(20000);
            sheet.Cells["B4"].PutValue(18000);
            sheet.Cells["B5"].PutValue(22000);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Use the cell value as the chart title
            chart.Title.Text = sheet.Cells["A1"].StringValue;
            chart.Title.IsVisible = true;

            // Freeze the row that contains the title text (row 1)
            sheet.PageSetup.PrintTitleRows = "$1:$1";

            // Save the workbook
            workbook.Save("ChartWithFrozenTitleRow.xlsx");
        }
    }
}
