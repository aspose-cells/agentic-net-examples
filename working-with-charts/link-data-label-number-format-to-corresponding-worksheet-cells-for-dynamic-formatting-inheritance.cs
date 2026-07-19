// Title: Aspose.Cells for .NET C# – Link Chart Data Labels to Worksheet Cells for Dynamic Number‑Format Inheritance
// Description: Demonstrates how to create a workbook, add category/value data, apply a custom currency format to a helper column, build a column chart, and configure the series so data labels pull their values from a linked range (C2:C3) and automatically inherit the cells' number format at runtime.
// Keywords: Aspose.Cells | C# chart data labels | LinkedSource | NumberFormatLinked | dynamic number format inheritance | Excel column chart | currency formatting in charts | Aspose.Cells .NET example | GitHub Aspose.Cells demo | chart label formatting
// Common Searches: Aspose.Cells link chart data label to cells | C# dynamic number format inheritance chart labels | Set DataLabels.LinkedSource in Aspose.Cells | NumberFormatLinked property usage Aspose.Cells | How to inherit cell format for chart labels .NET
// Developer Intent: Enable chart data labels to display values from a worksheet range and automatically adopt the range's number‑format without hard‑coding the format in code.
// Use Cases: Financial dashboards where chart labels must reflect currency formats defined in worksheet cells and update instantly when the format changes. | Locale‑aware reporting templates that let end‑users modify number formats in a helper column and see the changes reflected in chart labels. | Reusable chart components that pull both values and formatting from linked cells, reducing maintenance and eliminating duplicate format definitions.
// AI Prompts: Show how to change the linked source to a different range and apply a percentage format to the data labels. | Generate code that adjusts the LinkedSource range automatically based on the number of points in the series. | Explain the behavior and limitations of NumberFormatLinked in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add category/value data, apply a custom currency format to a helper column, build a column chart, and configure the series so data labels pull their values from a linked range (C2:C3) and automatically inherit the cells' number format at runtime.
    public class DataLabelsNumberFormatLinkedDynamicDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for categories and values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(1234.56);
                sheet.Cells["B3"].PutValue(7890.12);

                // Add cells that contain the formatted representation of the values
                sheet.Cells["C1"].PutValue("Formatted");
                sheet.Cells["C2"].PutValue("$1,234.56");
                sheet.Cells["C3"].PutValue("$7,890.12");

                // Apply a custom number format to the formatted cells (optional, demonstrates inheritance)
                Style style = workbook.CreateStyle();
                style.Custom = "$#,##0.00";
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;
                sheet.Cells.CreateRange("C2:C3").ApplyStyle(style, flag);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B3", true);          // Values
                chart.NSeries.CategoryData = "A2:A3";      // Categories

                // Configure data labels:
                // - Show the value
                // - Link the label source to the formatted cells
                // - Enable number format inheritance from the linked cells
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.LinkedSource = "C2:C3";
                series.DataLabels.NumberFormatLinked = true;

                // Save the workbook
                string outputPath = "DataLabelsNumberFormatLinkedDynamicDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsNumberFormatLinkedDynamicDemo.Run();
        }
    }
}
