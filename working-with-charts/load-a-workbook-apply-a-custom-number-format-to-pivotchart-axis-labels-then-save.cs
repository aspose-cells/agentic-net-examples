// Title: Aspose.Cells .NET – Apply Custom Number Format to PivotChart Axis Labels (C#)
// Description: Load a workbook, locate a PivotChart, set a custom number format for the value (and optionally category) axis tick labels, and save the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PivotChart number format | C# chart axis custom format | set value axis format Aspose.Cells | format pivot chart tick labels | apply currency format to chart axis | date format category axis Aspose.Cells | Aspose.Cells chart formatting example
// Common Searches: Aspose.Cells change PivotChart axis number format C# | how to set custom format for chart value axis using Aspose.Cells | apply dollar format to PivotChart axis labels .NET | date format for PivotChart category axis Aspose.Cells | save workbook after modifying chart axis Aspose.Cells
// Developer Intent: Modify the number format of a PivotChart’s axis tick labels and persist the changes in the workbook.
// Use Cases: Display monetary values on a PivotChart’s value axis with "$#,##0.00" formatting before exporting. | Show month‑day labels on the category axis of a PivotChart using "mmm dd" format. | Validate the presence of charts, apply distinct formats to each axis, and save the workbook in a single operation.
// AI Prompts: Generate a C# snippet that sets a custom number format for both value and category axes of a PivotChart with Aspose.Cells. | Provide code that iterates through all charts in a workbook and applies different axis formats based on chart type using Aspose.Cells for .NET. | Explain how to detect a PivotChart and safely apply axis number formatting with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace PivotChartNumberFormatExample
{
    // Load a workbook, locate a PivotChart, set a custom number format for the value (and optionally category) axis tick labels, and save the updated file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook that contains a PivotChart
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart (assumed to be a PivotChart)
            Chart chart = worksheet.Charts[0];

            // Apply a custom number format to the value axis tick labels
            // Example format: two decimal places with a dollar sign
            chart.ValueAxis.TickLabels.NumberFormat = "$#,##0.00";

            // Optionally, apply a custom format to the category axis tick labels
            // chart.CategoryAxis.TickLabels.NumberFormat = "mmm dd";

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("PivotChart axis labels formatted and workbook saved.");
        }
    }
}
