// Title: Aspose.Cells for .NET: Set Category Axis Tick Labels to Stacked (Vertical) Orientation
// Description: Demonstrates how to create a workbook, add a column chart, and apply ChartTextDirectionType.Stacked to display the category axis tick labels in a vertical (stacked) layout before saving the file.
// Keywords: Aspose.Cells | C# | .NET | ChartTextDirectionType | Stacked | vertical tick labels | category axis | chart label orientation | Excel chart formatting | example code
// Common Searches: Aspose.Cells set tick label direction stacked | ChartTextDirectionType.Stacked usage .NET | vertical category axis labels Aspose.Cells | how to rotate chart tick labels in C# | Aspose.Cells chart label orientation example
// Developer Intent: Apply ChartTextDirectionType.Stacked to rotate category axis tick labels vertically in an Aspose.Cells chart.
// Use Cases: Improve readability of long category names in column charts by stacking labels vertically. | Generate Excel reports where chart labels must fit narrow columns without truncation. | Automate consistent vertical label formatting across multiple charts in a workbook.
// AI Prompts: Show a C# example that sets chart tick labels to Stacked using Aspose.Cells. | Explain the effect of ChartTextDirectionType.Stacked on chart label orientation and how to implement it. | Provide step‑by‑step code to create a workbook, add a column chart, and apply vertical tick label direction.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTickLabelsStackedDemo
{
    // Demonstrates how to create a workbook, add a column chart, and apply ChartTextDirectionType.Stacked to display the category axis tick labels in a vertical (stacked) layout before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set tick labels direction to Stacked (vertical orientation)
            chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Stacked;

            // Save the workbook to a file
            workbook.Save("TickLabelsStackedDemo.xlsx");
        }
    }
}
