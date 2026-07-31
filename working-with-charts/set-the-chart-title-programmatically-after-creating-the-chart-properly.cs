// Title: Aspose.Cells for .NET – Programmatically Add a Title to a Column Chart (C#)
// Description: This C# example creates a new workbook, populates it with sample data, inserts a column chart, defines the series and categories, makes the chart title visible, sets the title text to "Sales Overview", and saves the file as ChartWithTitle.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells chart title C# | set chart title Aspose.Cells | column chart title .NET | chart title visibility Aspose.Cells | programmatic chart title | Aspose.Cells example | C# Excel chart title | Aspose.Cells Chart API | Excel automation chart title
// Common Searches: how to set chart title in Aspose.Cells C# | Aspose.Cells add title to column chart | make chart title visible Aspose.Cells | change chart title text programmatically Aspose.Cells | Aspose.Cells chart title example
// Developer Intent: Add a custom, visible title to a chart generated with Aspose.Cells in a .NET application.
// Use Cases: Automated sales reports where each chart needs a descriptive heading. | Dynamic dashboards that adjust chart titles based on user‑selected data. | Batch generation of Excel workbooks with multiple charts, ensuring every chart includes a clear title for presentation.
// AI Prompts: Provide C# code to set, update, or hide a chart title using Aspose.Cells. | Show how to change a chart title based on a variable value in Aspose.Cells for .NET. | Explain the steps to make a chart title visible and assign text in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleDemo
{
    // This C# example creates a new workbook, populates it with sample data, inserts a column chart, defines the series and categories, makes the chart title visible, sets the title text to "Sales Overview", and saves the file as ChartWithTitle.xlsx using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet (lifecycle rule: create chart)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title programmatically
            chart.Title.IsVisible = true;          // Ensure the title is visible
            chart.Title.Text = "Sales Overview";   // Set the desired title text

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChartWithTitle.xlsx");

            Console.WriteLine("Workbook saved with chart title set.");
        }
    }
}
