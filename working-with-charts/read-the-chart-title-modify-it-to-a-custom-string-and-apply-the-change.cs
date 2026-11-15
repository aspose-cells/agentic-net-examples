// Title: Set a custom title for an Excel chart using Aspose.Cells in C#
// Common Searches: C# code to change the title of an existing chart in an Excel file with Aspose.Cells | How to make a chart title visible and set its text using Aspose.Cells for .NET | Update the first chart title in a workbook programmatically with Aspose.Cells | Aspose.Cells example for reading and modifying chart title in .xlsx | Changing Excel chart title programmatically in C# using Aspose.Cells library
// Tags: Aspose.Cells set chart title | C# modify Excel chart title | Aspose.Cells chart title visibility | update chart title .NET | Excel chart title manipulation Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing .xlsx workbook, accesses the first worksheet, checks for charts, retrieves the first chart, assigns a custom string to its Title.Text, ensures the title is visible, and saves the workbook with the updated chart title.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the first chart in the worksheet
        // Ensure that the worksheet contains at least one chart
        if (sheet.Charts.Count > 0)
        {
            Chart chart = sheet.Charts[0];

            // Modify the chart title
            // If the chart does not have a title, this will create one
            chart.Title.Text = "My Custom Chart Title";

            // Optionally, make sure the title is visible
            chart.Title.IsVisible = true;
        }
        else
        {
            Console.WriteLine("No charts found in the worksheet.");
        }

        // Save the workbook with the updated chart title
        workbook.Save("output.xlsx");
    }
}
