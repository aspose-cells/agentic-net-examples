// Title: C# – Read and Set an Excel Chart Title with Aspose.Cells
// Description: Creates a new workbook, adds sample data, inserts a column chart, reads the current chart title, replaces it with a custom string, prints both titles to the console, and saves the file as ChartTitleModified.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart title C# | modify Excel chart title programmatically | read chart title Aspose.Cells | set custom chart title .NET | Aspose.Cells chart title example | C# Excel chart title update
// Common Searches: How to change a chart title in Excel using Aspose.Cells C# | Aspose.Cells read chart title C# example | Set custom title for column chart Aspose.Cells | Update Excel chart title programmatically .NET | Save workbook after modifying chart title Aspose.Cells
// Developer Intent: Read the existing chart title, replace it with a custom string, and save the workbook.
// Use Cases: Log the original chart title before modification for audit trails. | Apply a dynamic title derived from user input or calculated metrics. | Batch‑process multiple worksheets to enforce consistent branding across all chart titles.
// AI Prompts: Generate C# code that iterates through every chart in a workbook and sets each title to "Report Summary" with Aspose.Cells. | Show how to change a chart title only when its current text matches a specific value, using Aspose.Cells for .NET. | Explain best practices for localizing Excel chart titles into different languages when creating files with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleExample
{
    // Creates a new workbook, adds sample data, inserts a column chart, reads the current chart title, replaces it with a custom string, prints both titles to the console, and saves the file as ChartTitleModified.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set an initial title (optional)
            chart.Title.Text = "Original Chart Title";

            // Read the current chart title
            string currentTitle = chart.Title.Text;
            Console.WriteLine("Current Chart Title: " + currentTitle);

            // Modify the chart title to a custom string
            chart.Title.Text = "Custom Chart Title";

            // Verify the change
            Console.WriteLine("Updated Chart Title: " + chart.Title.Text);

            // Save the workbook (lifecycle: save)
            workbook.Save("ChartTitleModified.xlsx");
        }
    }
}
