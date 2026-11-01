// Title: Create a column chart with a cell‑derived title and freeze the title row using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that inserts a column chart, reads the title text from cell A1, assigns it to the chart, makes the title visible, and freezes row 1 with Aspose.Cells. | Generate a .NET snippet that binds a worksheet cell value to a chart title, defines the series range, applies FreezePanes to keep the header row static, and saves the workbook. | Provide an Aspose.Cells example that updates an existing workbook to use a cell value as the chart title and locks the first row in place while scrolling.
// Common Searches: aspocells set chart title from a worksheet cell c# | how to freeze the header row in an Excel file using Aspose.Cells .NET | c# create column chart with dynamic title from cell using Aspose.Cells | aspocells freeze panes first row while scrolling chart
// Tags: use cell value as chart title Aspose.Cells C# | freeze panes on first row Aspose.Cells | create column chart Aspose.Cells .NET | bind worksheet cell to chart title Aspose.Cells | Aspose.Cells freeze row example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a new workbook, writes "Sales Report" to cell A1, adds a column chart, sets the chart title to the cell's text, defines series and category data, freezes the first row so the title stays visible while scrolling, and saves the file as ChartWithTitleAndFrozenRow.xlsx.
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

            // Put title text in cell A1
            sheet.Cells["A1"].PutValue("Sales Report");

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title using the text from cell A1
            chart.Title.Text = sheet.Cells["A1"].StringValue;
            chart.Title.IsVisible = true;

            // Sample data for the chart
            sheet.Cells["B2"].PutValue("Q1");
            sheet.Cells["C2"].PutValue(120);
            sheet.Cells["B3"].PutValue("Q2");
            sheet.Cells["C3"].PutValue(150);
            sheet.Cells["B4"].PutValue("Q3");
            sheet.Cells["C4"].PutValue(130);
            sheet.Cells["B5"].PutValue("Q4");
            sheet.Cells["C5"].PutValue(170);

            // Define the series and category data ranges
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries.CategoryData = "B2:B5";

            // Freeze the first row (where the title resides) so it stays visible while scrolling
            sheet.FreezePanes(1, 0, 1, 0); // Freeze rows above row 2 (i.e., row 1)

            // Determine output file path
            string outputPath = "ChartWithTitleAndFrozenRow.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
