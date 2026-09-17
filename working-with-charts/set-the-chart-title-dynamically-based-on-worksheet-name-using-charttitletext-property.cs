// Title: Set an Excel column chart title to the worksheet name at runtime using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add a column chart, and assign the chart's Title.Text property to the current worksheet's Name with Aspose.Cells in C#. | Write C# code that generates an Excel file where the chart title automatically reflects the sheet's name after renaming the worksheet. | Show how to bind a chart title to a worksheet name by setting Chart.Title.Text after inserting a column chart using Aspose.Cells.
// Common Searches: how to set chart title to sheet name using Aspose.Cells C# | Aspose.Cells Chart.Title.Text example for dynamic titles | C# create column chart and use worksheet name as chart title in Excel | programmatically change Excel chart title based on worksheet name Aspose.Cells | dynamic chart title from worksheet name Aspose.Cells C#
// Tags: Aspose.Cells chart title from worksheet name | C# set Chart.Title.Text property | create column chart Aspose.Cells | dynamic Excel chart title | bind chart title to sheet name

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new workbook, renames the first worksheet to "SalesData", adds a column chart, sets the chart's title to the worksheet's name via the Chart.Title.Text property, and saves the workbook as ChartWithDynamicTitle.xlsx, with exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Optionally set a custom name for the worksheet
            sheet.Name = "SalesData";

            // Add a column chart to the worksheet (lifecycle create)
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title dynamically based on the worksheet name
            chart.Title.Text = sheet.Name;

            // Save the workbook to a file (lifecycle save)
            workbook.Save("ChartWithDynamicTitle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
