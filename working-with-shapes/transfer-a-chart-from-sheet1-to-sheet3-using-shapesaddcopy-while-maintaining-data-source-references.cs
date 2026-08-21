// Title: Aspose.Cells for .NET – Copy a Chart from Sheet1 to Sheet3 with Shapes.AddCopy (preserve data source)
// Description: Demonstrates how to create a workbook, add sample data, build a column chart on Sheet1, retrieve its ChartShape, and duplicate the chart onto Sheet3 using Shapes.AddCopy while keeping the original data references intact. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells copy chart | Shapes.AddCopy example | chart shape duplication .NET | preserve chart data source | transfer chart between worksheets | C# Aspose.Cells chart copy | Excel chart cloning Aspose
// Common Searches: copy chart to another sheet Aspose.Cells C# | Shapes.AddCopy keep chart data link | duplicate Excel chart programmatically | move chart between worksheets Aspose | how to clone chart shape in Aspose.Cells
// Developer Intent: Programmatically duplicate an existing chart on a different worksheet without breaking its data bindings.
// Use Cases: Create a dashboard sheet that aggregates charts from multiple data sheets. | Generate summary reports where the same chart appears on a cover or overview page. | Automate workbook layouts that require identical charts on several worksheets.
// AI Prompts: Generate C# code using Aspose.Cells to copy a chart from Sheet1 to Sheet3 with Shapes.AddCopy while retaining the original data range. | Explain each parameter of Shapes.AddCopy when copying a chart shape and how they control the chart's placement on the target sheet. | Suggest an alternative method to copy a chart to another worksheet and adjust its top/left offsets for precise positioning.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add sample data, build a column chart on Sheet1, retrieve its ChartShape, and duplicate the chart onto Sheet3 using Shapes.AddCopy while keeping the original data references intact. The workbook is then saved as an Excel file.
class TransferChartExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (Sheet1) and rename it
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a third worksheet (Sheet3) where the chart will be copied
        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

        // Populate sample data in Sheet1 for the chart
        sheet1.Cells["A1"].PutValue("Category");
        sheet1.Cells["B1"].PutValue("Value");
        sheet1.Cells["A2"].PutValue("A");
        sheet1.Cells["B2"].PutValue(10);
        sheet1.Cells["A3"].PutValue("B");
        sheet1.Cells["B3"].PutValue(20);
        sheet1.Cells["A4"].PutValue("C");
        sheet1.Cells["B4"].PutValue(30);

        // Add a column chart to Sheet1
        int chartIndex = sheet1.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet1.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Retrieve the chart shape (ChartShape derives from Shape)
        Shape chartShape = chart.ChartObject;

        // Copy the chart shape to Sheet3 using Shapes.AddCopy
        // Parameters: source shape, top row, top offset (pixels), left column, left offset (pixels)
        sheet3.Shapes.AddCopy(chartShape, 5, 0, 0, 0);

        // Save the workbook
        workbook.Save("ChartCopyDemo.xlsx");
    }
}
