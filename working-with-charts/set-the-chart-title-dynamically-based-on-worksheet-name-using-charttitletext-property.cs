// Title: Aspose.Cells C# – Dynamically Assign Worksheet Name to Chart Title
// Description: This .NET example shows how to create a workbook with Aspose.Cells, rename the first worksheet, fill it with sample sales data, add a column chart, and programmatically set the chart's title to the worksheet's name using the Chart.Title.Text property. The title is made visible and the file is saved as DynamicChartTitle.xlsx.
// Keywords: Aspose.Cells | C# chart title | Chart.Title.Text | dynamic worksheet title | Excel chart automation | column chart Aspose | set chart title programmatically | Aspose.Cells .NET | Excel workbook generation | chart title visibility
// Common Searches: Aspose.Cells set chart title from sheet name C# | Chart.Title.Text example Aspose.Cells | How to make chart title visible in Aspose.Cells | Dynamic chart titles in Excel using Aspose.Cells | C# code to assign worksheet name to chart title
// Developer Intent: Assign the chart title to the current worksheet's name at runtime.
// Use Cases: Automated sales reports where each sheet’s chart displays its sheet name. | Multi‑sheet workbooks with matching chart titles for departmental dashboards. | Template‑driven Excel generation that avoids hard‑coded titles. | Batch processing of workbooks to ensure chart titles reflect worksheet identifiers.
// AI Prompts: Generate C# Aspose.Cells code that sets a chart title to the active worksheet name and ensures the title is visible. | Show a loop that iterates through all worksheets and updates each chart’s title with the corresponding sheet name. | Explain how to customize the font size, color, and alignment of a dynamically set chart title using Aspose.Cells. | Provide guidance on exporting the workbook with dynamic chart titles to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartTitle
{
    // This .NET example shows how to create a workbook with Aspose.Cells, rename the first worksheet, fill it with sample sales data, add a column chart, and programmatically set the chart's title to the worksheet's name using the Chart.Title.Text property. The title is made visible and the file is saved as DynamicChartTitle.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optionally set a custom name for the worksheet
            worksheet.Name = "SalesData";

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");
            worksheet.Cells["B1"].PutValue("Revenue");
            worksheet.Cells["B2"].PutValue(15000);
            worksheet.Cells["B3"].PutValue(18000);
            worksheet.Cells["B4"].PutValue(21000);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Dynamically set the chart title to the worksheet's name
            chart.Title.Text = worksheet.Name;
            chart.Title.IsVisible = true; // Ensure the title is displayed

            // Save the workbook (lifecycle: save)
            workbook.Save("DynamicChartTitle.xlsx");
        }
    }
}
