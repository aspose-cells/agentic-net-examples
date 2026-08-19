// Title: Aspose.Cells C# – Set Chart Title Dynamically from Worksheet Name
// Description: Demonstrates how to create a workbook, rename a worksheet, add sample data, insert a column chart, and assign the chart title to the worksheet's Name property using Chart.Title.Text in Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart title | dynamic chart title C# | Chart.Title.Text example | set chart title from sheet name | Aspose.Cells column chart | programmatic chart title | C# Aspose.Cells tutorial
// Common Searches: Aspose.Cells set chart title from worksheet name | C# dynamic chart title Aspose.Cells | Chart.Title.Text usage Aspose.Cells .NET | how to make chart title visible Aspose.Cells | assign worksheet name to chart title Aspose.Cells
// Developer Intent: Programmatically set a chart's title to match the current worksheet's name using Aspose.Cells for .NET.
// Use Cases: Automated reports where each sheet’s chart reflects the sheet topic without manual edits. | Templates with identical chart layouts that adapt titles when worksheets are renamed or added. | Batch processing of multiple worksheets to ensure chart titles stay in sync with sheet names.
// AI Prompts: Show C# code that loops through all worksheets in an Aspose.Cells workbook and sets each chart's title to the worksheet name. | Provide an example of creating a line chart in Aspose.Cells and assigning its title from the worksheet name, ensuring the title is visible. | Explain how to conditionally hide a chart title in Aspose.Cells when the worksheet name is empty or null.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartTitle
{
    // Demonstrates how to create a workbook, rename a worksheet, add sample data, insert a column chart, and assign the chart title to the worksheet's Name property using Chart.Title.Text in Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optionally rename the worksheet to demonstrate dynamic title
            worksheet.Name = "SalesData";

            // Add sample data for the chart
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");
            worksheet.Cells["B1"].PutValue("Revenue");
            worksheet.Cells["B2"].PutValue(15000);
            worksheet.Cells["B3"].PutValue(18000);
            worksheet.Cells["B4"].PutValue(21000);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Dynamically set the chart title based on the worksheet name
            chart.Title.Text = worksheet.Name;
            chart.Title.IsVisible = true; // Ensure the title is displayed

            // Save the workbook to a file
            workbook.Save("DynamicChartTitle.xlsx");
        }
    }
}
