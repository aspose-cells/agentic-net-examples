// Title: Set a chart title to the worksheet name dynamically with Aspose.Cells for .NET (C#)
// Common Searches: how to set chart title to worksheet name using Aspose.Cells in C# | Aspose.Cells C# assign sheet name as chart title | make chart title visible when using Aspose.Cells chart.Title | create column chart and set its title from the sheet name with Aspose.Cells .NET | programmatically set Chart.Title.Text based on worksheet name Aspose.Cells
// Tags: chart.Title.Text property Aspose.Cells | dynamic chart title based on worksheet name | chart title visibility setting Aspose.Cells | column chart creation Aspose.Cells C# | assign worksheet name to chart title .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new workbook, renames the first worksheet to "SalesData", adds a column chart, sets the chart's title to the worksheet's name via the Chart.Title.Text property, ensures the title is visible, and saves the file as ChartWithDynamicTitle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // Set a custom name for demonstration; in real scenarios the name may already be set
            sheet.Name = "SalesData";

            // Add a column chart to the worksheet (lifecycle rule: create)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title dynamically based on the worksheet name
            chart.Title.Text = sheet.Name;
            chart.Title.IsVisible = true; // Ensure the title is displayed

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChartWithDynamicTitle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
