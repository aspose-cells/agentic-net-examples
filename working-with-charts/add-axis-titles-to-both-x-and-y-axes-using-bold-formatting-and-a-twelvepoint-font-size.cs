// Title: Set bold 12‑pt X‑ and Y‑axis titles in an Aspose.Cells .NET chart
// Description: Creates a workbook, adds sample data and a column chart, then makes the Category (X) and Value (Y) axis titles visible, sets their text, applies bold styling and a 12‑point font, and saves the file.
// Keywords: Aspose.Cells chart axis title C# | set axis font size Aspose.Cells | .NET chart axis formatting | bold axis titles Excel library | add X axis title Aspose.Cells | add Y axis title Aspose.Cells | chart axis title visibility | Aspose.Cells column chart example | C# Excel chart customization | Aspose.Cells axis title font
// Common Searches: How to add bold axis titles in Aspose.Cells C# | Set 12 point font for X and Y axis titles in .NET chart | Make chart axis titles visible with Aspose.Cells | Formatting axis titles in an Aspose.Cells column chart | Aspose.Cells example for axis title styling
// Developer Intent: Apply visible, bold, 12‑pt titles to both the X (category) and Y (value) axes of a chart using Aspose.Cells for .NET.
// Use Cases: Generate a sales performance workbook where the chart axes are clearly labeled for presentation to stakeholders. | Build a financial dashboard Excel file with column charts that emphasize axis titles for quick data interpretation. | Export analytical results to Excel with professionally formatted charts that include bold axis headings.
// AI Prompts: Show C# code to set the text, visibility, bold style, and 12‑pt font for X and Y axis titles in an Aspose.Cells chart. | Provide an Aspose.Cells example that changes the axis title font family, size, and style after the chart is created. | Explain how to toggle axis title visibility and apply custom formatting (color, alignment) programmatically with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisTitleExample
{
    // Creates a workbook, adds sample data and a column chart, then makes the Category (X) and Value (Y) axis titles visible, sets their text, applies bold styling and a 12‑point font, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure X (Category) axis title
            chart.CategoryAxis.Title.Text = "Categories";
            chart.CategoryAxis.Title.IsVisible = true;
            chart.CategoryAxis.Title.Font.IsBold = true;
            chart.CategoryAxis.Title.Font.Size = 12;

            // Configure Y (Value) axis title
            chart.ValueAxis.Title.Text = "Values";
            chart.ValueAxis.Title.IsVisible = true;
            chart.ValueAxis.Title.Font.IsBold = true;
            chart.ValueAxis.Title.Font.Size = 12;

            // Save the workbook
            workbook.Save("AxisTitles_Output.xlsx");
        }
    }
}
