// Title: Enable Text Wrapping for Pie Chart Data Labels in Aspose.Cells (C#/.NET)
// Description: C# example that creates an Excel workbook, adds a pie chart with values and category names, and activates text wrapping for the data labels by setting DataLabels.IsTextWrapped = true. The file is saved as PieChart_With_WrappedDataLabels.xlsx.
// Keywords: Aspose.Cells | C# | .NET | pie chart | data labels | text wrapping | IsTextWrapped | chart label wrap | Excel generation | sample code | GitHub example | Aspose.Cells chart API
// Common Searches: Aspose.Cells wrap pie chart data labels | DataLabels.IsTextWrapped C# example | how to enable text wrap for chart labels in Aspose.Cells | C# pie chart label wrapping Aspose | Aspose.Cells chart data label options
// Developer Intent: Wrap the text of pie‑chart data labels using Aspose.Cells.
// Use Cases: Generate Excel reports where long category names in a pie chart stay readable without truncation. | Build dashboards that display both numeric values and wrapped category names on pie‑chart labels. | Automate workbook creation for presentations, ensuring chart annotations fit within label boxes. | Provide a reusable snippet for developers needing chart‑label formatting in Aspose.Cells.
// AI Prompts: Show C# code to set DataLabels.IsTextWrapped = true for a pie chart in Aspose.Cells. | Explain how text wrapping changes the appearance of pie‑chart data labels in an Excel file generated with Aspose.Cells. | Give a step‑by‑step guide to add a pie chart with wrapped data labels using Aspose.Cells for .NET. | Provide a GitHub‑style snippet that creates a workbook, adds a pie chart, and enables label wrapping.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPieChartDataLabelWrap
{
    // C# example that creates an Excel workbook, adds a pie chart with values and category names, and activates text wrapping for the data labels by setting DataLabels.IsTextWrapped = true. The file is saved as PieChart_With_WrappedDataLabels.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
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
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            pieChart.NSeries.Add("B2:B4", true);
            pieChart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels labels = pieChart.NSeries[0].DataLabels;
            labels.ShowValue = true;          // Show the numeric values
            labels.ShowCategoryName = true;   // Show the category names

            // Enable text wrapping for the data labels
            labels.IsTextWrapped = true;

            // Save the workbook to a file
            workbook.Save("PieChart_With_WrappedDataLabels.xlsx");
        }
    }
}
