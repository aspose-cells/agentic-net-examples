// Title: Detect a Chart’s X Axis Type and Convert to a Numeric Value Axis with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all charts in a workbook, verifies whether the CategoryAxis is non‑numeric, switches the chart to a Scatter type, and assigns a numeric cell range to the series XValues. | Create a method that loads an Excel file, determines if each chart’s X axis is a value axis, and if not, programmatically changes the axis to numeric by updating the chart type and providing X and Y data ranges.
// Common Searches: Aspose.Cells determine chart X axis type programmatically | convert Excel chart category axis to numeric axis using Aspose.Cells | change chart to scatter to enable numeric X values in Aspose.Cells | set XValues for chart series with Aspose.Cells C# | replace non‑value X axis in Excel charts via Aspose.Cells
// Tags: chart category axis detection Aspose.Cells | convert chart to scatter Aspose.Cells | assign numeric XValues range Aspose.Cells C# | programmatic Excel chart axis conversion .NET | numeric X axis for Excel chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, loops through each chart on the first worksheet, checks whether the CategoryAxis is a value (numeric) axis, and if it is not, changes the chart type to Scatter, populates cells with numeric data, assigns those ranges to the series XValues and Values, and saves the updated workbook.
class DetectAndConvertChartXAxis
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Work with the first worksheet (adjust as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Iterate through all charts in the worksheet
        foreach (Chart chart in sheet.Charts)
        {
            // The X axis of a chart is the CategoryAxis
            Axis xAxis = chart.CategoryAxis;

            // Determine if the X axis is currently a value‑type axis.
            // In Aspose.Cells a CategoryScale indicates a textual/category axis,
            // while DateScale or AutomaticScale can be used for numeric/date data.
            bool isValueAxis = xAxis.CategoryType != CategoryType.CategoryScale;

            if (!isValueAxis)
            {
                // X axis is not a value axis – convert it to support numeric data.

                // 1. Change the chart type to Scatter (X axis is a true value axis in scatter charts)
                chart.Type = ChartType.Scatter;

                // 2. Ensure there is a numeric range in the worksheet to use as X values.
                //    For demonstration we fill cells A2:A5 with numbers 1‑4.
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1); // 1,2,3,4
                }

                // 3. Assign the numeric range to the series' XValues property.
                //    This makes the X axis display numeric data.
                if (chart.NSeries.Count > 0)
                {
                    Series series = chart.NSeries[0];
                    series.XValues = "A2:A5";

                    // Optional: also set Y values if they are not already set.
                    // Here we use B2:B5 as sample Y values.
                    for (int i = 2; i <= 5; i++)
                    {
                        sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // 10,20,30,40
                    }
                    series.Values = "B2:B5";
                }
            }
            else
            {
                // X axis is already a value axis – no action needed.
                Console.WriteLine("Chart already has a numeric X axis.");
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
