// Title: Add identical column charts to multiple worksheets with Aspose.Cells for .NET
// Description: Shows how to create a workbook, ensure a specific number of worksheets, populate each sheet with its own data, insert a column chart that uses the same range on every sheet, assign a sheet‑specific title, and save the result as BatchCharts.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | .NET | column chart | multiple worksheets | batch chart creation | programmatic chart insertion | chart title per sheet | Excel automation | Workbook.Save | ChartType.Column
// Common Searches: Aspose.Cells add same chart to each worksheet | C# create column chart on multiple sheets | batch insert charts Aspose.Cells .NET | loop to add charts to worksheets | set chart title per sheet Aspose.Cells
// Developer Intent: Insert the same chart type into several worksheets while applying sheet‑specific data and titles.
// Use Cases: Regional sales workbook where each sheet displays a column chart of that region’s figures. | Quarterly financial dashboard that adds identical performance charts to each quarter’s worksheet. | Template generator that programmatically adds a chart to any number of sheets based on dynamic data.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line chart to every worksheet, using a unique data range per sheet and custom titles. | Show how to create five worksheets, fill them with sample data, and place a pie chart on each sheet with distinct titles and positions. | Explain how to control chart size and position programmatically when adding charts in a batch loop with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, ensure a specific number of worksheets, populate each sheet with its own data, insert a column chart that uses the same range on every sheet, assign a sheet‑specific title, and save the result as BatchCharts.xlsx using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();
            WorksheetCollection worksheets = workbook.Worksheets;

            // Desired number of worksheets
            int sheetCount = 3;

            // Ensure the workbook has enough worksheets
            while (worksheets.Count < sheetCount)
            {
                worksheets.Add();
            }

            for (int i = 0; i < sheetCount; i++)
            {
                // Access the worksheet (already exists)
                Worksheet sheet = worksheets[i];
                sheet.Name = $"Sheet{i + 1}";

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int row = 2; row <= 5; row++)
                {
                    sheet.Cells[$"A{row}"].PutValue($"Item{row - 1}");
                    // Distinct values per sheet
                    sheet.Cells[$"B{row}"].PutValue((i + 1) * row * 10);
                }

                // Add a column chart with a specific data range
                int chartIdx = sheet.Charts.Add(
                    ChartType.Column,   // Chart type
                    "A1:B5",            // Data range
                    true,               // Plot by column
                    6, 0,               // Upper‑left cell (row, column)
                    20, 8);             // Lower‑right cell (row, column)

                Chart chart = sheet.Charts[chartIdx];
                chart.Title.Text = $"Chart for {sheet.Name}";
            }

            // Save the workbook
            workbook.Save("BatchCharts.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
