// Title: Apply 18‑pt Dark Blue Title to an Aspose.Cells Column Chart (C#)
// Description: Creates a new workbook, inserts month‑sales data, adds a column chart, makes the title visible, sets its text to “Quarterly Sales Report”, and formats the title with an 18‑point dark‑blue font before saving as ChartWithCustomTitle.xlsx.
// Keywords: Aspose.Cells | C# | .NET | column chart | chart title | font size | font color | 18pt | dark blue | Excel automation | custom chart title
// Common Searches: Aspose.Cells set chart title font size C# | change chart title color Aspose.Cells .NET | add visible chart title column chart Aspose.Cells | customize chart title appearance Aspose.Cells | format chart title in Excel using Aspose.Cells
// Developer Intent: Add a styled title to a chart in an Excel workbook.
// Use Cases: Standardized sales reports with branded chart titles | Automated generation of presentation‑ready Excel charts | Dashboards that require consistent title styling across multiple workbooks | Compliance with corporate visual guidelines for exported Excel files
// AI Prompts: Generate C# code with Aspose.Cells to create a line chart and set the title to 20pt red font. | Show how to change the chart title font family and alignment in Aspose.Cells .NET. | Explain how to toggle chart title visibility based on data conditions using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleDemo
{
    // Creates a new workbook, inserts month‑sales data, adds a column chart, makes the title visible, sets its text to “Quarterly Sales Report”, and formats the title with an 18‑point dark‑blue font before saving as ChartWithCustomTitle.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the chart title with custom font size and color
            chart.Title.Text = "Quarterly Sales Report";
            chart.Title.IsVisible = true;                     // Ensure the title is displayed
            chart.Title.Font.Size = 18;                       // Custom font size
            chart.Title.Font.Color = Color.DarkBlue;          // Custom font color

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChartWithCustomTitle.xlsx");
        }
    }
}
