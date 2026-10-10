// Title: Set a column chart title to the worksheet name dynamically with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a column chart in a new workbook, renames the worksheet, and assigns the worksheet's name to the chart's Title.Text while making the title visible. | Write a C# snippet using Aspose.Cells to add a column chart, set its title based on the active sheet name, and ensure the title appears in the saved Excel file.
// Common Searches: Aspose.Cells C# set chart title to current worksheet name | how to bind column chart title to sheet name using Aspose.Cells .NET | make chart title visible after assigning Title.Text in Aspose.Cells
// Tags: Aspose.Cells set chart title programmatically | column chart title from worksheet name .NET | Aspose.Cells chart title visibility | C# add column chart Aspose.Cells | dynamic chart title based on sheet name

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// // Creates a workbook, renames the first worksheet to "SalesData", adds a column chart, sets the chart title to the worksheet name, makes the title visible, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SalesData";

            // Add a column chart to the worksheet (position: row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title dynamically based on the worksheet name
            chart.Title.Text = sheet.Name;
            chart.Title.IsVisible = true; // Ensure the title is displayed

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
