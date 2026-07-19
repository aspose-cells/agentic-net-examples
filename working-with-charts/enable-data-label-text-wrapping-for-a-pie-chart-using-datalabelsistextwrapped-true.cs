// Title: Wrap Pie Chart Data Labels with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a pie chart, enable data labels, and set DataLabels.IsTextWrapped = true so label text automatically wraps onto multiple lines. The example saves the result as an Excel file.
// Keywords: Aspose.Cells C# pie chart | DataLabels.IsTextWrapped | chart label wrapping .NET | wrap pie chart data labels | Aspose.Cells chart formatting
// Common Searches: Aspose.Cells wrap text in chart data labels | C# set DataLabels.IsTextWrapped for pie chart | How to enable multiline labels in Aspose.Cells chart | Aspose.Cells chart label wrap example
// Developer Intent: Apply text wrapping to data labels on a pie chart using Aspose.Cells.
// Use Cases: Display long category names without truncation in pie charts. | Produce printable Excel reports where labels need multiple lines. | Automate Excel generation with readable, wrapped chart labels.
// AI Prompts: Show C# code that creates a pie chart with wrapped data labels using Aspose.Cells. | Explain the effect of DataLabels.IsTextWrapped together with ShowCategoryName and ShowValue. | Provide a step‑by‑step guide to enable multiline data labels in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelWrapDemo
{
    // Demonstrates how to create a workbook, add a pie chart, enable data labels, and set DataLabels.IsTextWrapped = true so label text automatically wraps onto multiple lines. The example saves the result as an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["B4"].PutValue(65);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;          // Show the numeric values
            dataLabels.ShowCategoryName = true;   // Show the category names

            // Enable text wrapping for the data labels
            dataLabels.IsTextWrapped = true;

            // Save the workbook to a file
            workbook.Save("PieChart_DataLabels_Wrapped.xlsx");
        }
    }
}
