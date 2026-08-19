// Title: Aspose.Cells for .NET – C# – Set secondary axis title “Revenue (USD)” with bold‑italic font
// Description: This example demonstrates how to create a workbook, add a column chart with primary (Sales) and secondary (Revenue) series, enable the secondary axis, make its title visible, set the text to "Revenue (USD)", and apply bold and italic styling (including optional color) before saving the file.
// Keywords: Aspose.Cells | .NET | C# | chart secondary axis | secondary value axis title | bold italic font | Revenue USD | dual axis chart | Excel chart formatting | Aspose.Cells example
// Common Searches: Aspose.Cells set secondary axis title C# | C# chart secondary value axis bold italic | how to format secondary axis label Aspose.Cells | dual axis column chart Aspose.Cells .NET example | change secondary axis title color Aspose.Cells
// Developer Intent: Add a visible secondary value‑axis title "Revenue (USD)" to a chart and format the title font as bold and italic using Aspose.Cells for .NET.
// Use Cases: Financial dashboards that display revenue on a secondary axis with a highlighted title. | Automated report generation where dual‑axis charts need distinct styling for clarity. | Excel workbooks prepared for presentations, requiring bold‑italic secondary axis labels.
// AI Prompts: Generate C# code with Aspose.Cells to create a column chart, enable a secondary axis, set its title to "Revenue (USD)", and apply bold‑italic formatting. | Show how to style the secondary value axis title (visibility, text, bold, italic, color) in an Aspose.Cells chart for .NET. | Explain the steps to add a dual‑axis chart in Aspose.Cells and customize the secondary axis label for emphasis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsSecondaryAxisDemo
{
    // This example demonstrates how to create a workbook, add a column chart with primary (Sales) and secondary (Revenue) series, enable the secondary axis, make its title visible, set the text to "Revenue (USD)", and apply bold and italic styling (including optional color) before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1500);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(1800);

            sheet.Cells["C1"].PutValue("Revenue");
            sheet.Cells["C2"].PutValue(30000);
            sheet.Cells["C3"].PutValue(45000);
            sheet.Cells["C4"].PutValue(40000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // First series (Sales) on primary axis
            chart.NSeries.Add("B2:B4", true);
            // Second series (Revenue) on secondary axis
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].PlotOnSecondAxis = true; // Enable secondary axis for the second series

            // Set category (X) data
            chart.NSeries.CategoryData = "A2:A4";

            // Configure secondary value axis title
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.IsVisible = true;
            secondaryAxis.Title.Text = "Revenue (USD)";
            secondaryAxis.Title.Font.IsBold = true;
            secondaryAxis.Title.Font.IsItalic = true;
            secondaryAxis.Title.Font.Color = Color.DarkGreen; // optional styling

            // (Optional) Make primary axis titles visible for completeness
            chart.ValueAxis.Title.IsVisible = true;
            chart.ValueAxis.Title.Text = "Sales (Units)";
            chart.CategoryAxis.Title.IsVisible = true;
            chart.CategoryAxis.Title.Text = "Quarter";

            // Save the workbook
            workbook.Save("SecondaryAxisTitleDemo.xlsx");
        }
    }
}
