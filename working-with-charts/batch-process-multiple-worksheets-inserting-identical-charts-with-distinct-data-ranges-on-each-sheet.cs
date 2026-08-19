// Title: Add identical column charts to multiple worksheets with sheet‑specific data using Aspose.Cells for .NET (C#)
// Description: C# example that creates a new workbook, removes the default sheet, loops to add several worksheets, fills each with its own data set, inserts a column chart at a fixed position, sets the range to A1:B6, gives the chart a title that includes the sheet name, shows the legend, and saves the file as BatchChartsOutput.xlsx using the Aspose.Cells API.
// Keywords: Aspose.Cells C# add chart | multiple worksheets chart Aspose.Cells | batch create charts .NET | column chart programmatically Aspose | loop worksheets Aspose.Cells | set chart data range per sheet | Aspose.Cells chart example | Excel automation C# Aspose | chart title with sheet name | save workbook Aspose.Cells
// Common Searches: add same chart to each worksheet Aspose.Cells | C# loop to create worksheets and insert column charts | Aspose.Cells set chart data range for each sheet | save workbook after adding charts Aspose.Cells | batch chart creation Aspose.Cells .NET | how to generate charts on multiple sheets using Aspose.Cells
// Developer Intent: Insert a column chart with identical layout into several worksheets, each using its own data range.
// Use Cases: Produce a regional sales report where each region gets a separate sheet with a column chart reflecting its figures. | Create a template that automatically adds a performance chart to every new worksheet, pulling data from that sheet. | Automate monthly KPI dashboards by looping through sheets, placing identical charts, and customizing titles per sheet.
// AI Prompts: Write C# code with Aspose.Cells that adds a line chart to every worksheet in a workbook, using data from A1:C10 on each sheet and setting the chart title to include the sheet name. | Provide an example that creates five worksheets, fills each with random numeric data, and inserts a pie chart on each sheet with a legend and a title that references the worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchChartDemo
{
    // C# example that creates a new workbook, removes the default sheet, loops to add several worksheets, fills each with its own data set, inserts a column chart at a fixed position, sets the range to A1:B6, gives the chart a title that includes the sheet name, shows the legend, and saves the file as BatchChartsOutput.xlsx using the Aspose.Cells API.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and remove the default worksheet
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();

                // Define number of worksheets to process
                int sheetCount = 3;

                // Loop to add worksheets and populate distinct data
                for (int i = 0; i < sheetCount; i++)
                {
                    // Add a new worksheet with a unique name
                    string sheetName = $"Sheet{i + 1}";
                    Worksheet sheet = workbook.Worksheets.Add(sheetName);

                    // Populate sample data: categories in column A, values in column B
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");
                    for (int row = 2; row <= 6; row++)
                    {
                        sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                        // Use different values for each sheet to make the chart distinct
                        sheet.Cells[$"B{row}"].PutValue((row - 1) * (i + 1) * 10);
                    }

                    // Add a column chart to the worksheet at a fixed position
                    // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 7);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set the data range for the chart (specific to the current sheet)
                    // The range "A1:B6" includes the header and data rows
                    chart.NSeries.Add("A1:B6", true);
                    chart.Title.Text = $"Sample Chart for {sheetName}";
                    chart.ShowLegend = true;
                }

                // Save the workbook to a file
                workbook.Save("BatchChartsOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
