// Title: C# – Add a Secondary Axis to a Line Series in a Combo Chart with Aspose.Cells
// Description: Creates a workbook, inserts sales and profit data, builds a column‑line combo chart, plots the profit line on a secondary Y‑axis, customizes the axis title, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells combo chart secondary axis | C# plot line series on secondary axis | Aspose.Cells column and line chart | secondary Y axis title Aspose.Cells | Excel combo chart with secondary axis .NET
// Common Searches: Aspose.Cells add secondary axis to line series | C# combo chart secondary Y axis example | plot line series on secondary axis Aspose.Cells | customize secondary axis title Aspose.Cells chart | create column‑line combo chart .NET
// Developer Intent: Add a line series to a combo chart and display it on a secondary Y‑axis for clearer data comparison.
// Use Cases: Compare monthly sales (columns) with profit margins (line) when the two metrics have different scales. | Generate a single Excel chart that visualizes revenue and cost‑to‑serve, using a secondary axis for the cost line. | Label the secondary axis to indicate the metric represented by the line series, improving report readability.
// AI Prompts: Show how to set a custom number format and marker style for the secondary axis line series in Aspose.Cells C#. | Explain step‑by‑step how to create a combo chart with multiple secondary axes using Aspose.Cells for .NET. | Generate code that adds data labels to the secondary line series and changes the secondary axis title font color.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartSecondaryAxis
{
    // Creates a workbook, inserts sales and profit data, builds a column‑line combo chart, plots the profit line on a secondary Y‑axis, customizes the axis title, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Primary series (Column)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Secondary series (Line)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(55);
            sheet.Cells["C5"].PutValue(70);

            // Add a combo chart (initially a Column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the primary (column) series
            chart.NSeries.Add("B2:B5", true);
            // Add the secondary (line) series
            chart.NSeries.Add("C2:C5", true);

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Change the second series type to Line to create a combo chart
            chart.NSeries[1].Type = ChartType.Line;

            // Plot the line series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize the secondary axis (e.g., title)
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Profit (Secondary Axis)";
            secondaryAxis.Title.Font.IsBold = true;

            // Save the workbook
            workbook.Save("ComboChartWithSecondaryAxis.xlsx");
        }
    }
}
