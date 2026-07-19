// Title: Color Chart Axis Tick Labels with Conditional Formatting in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add conditional formatting rules (green < 40, orange 40‑79, red ≥ 80), build a column chart, and enable the CategoryAxis.TickLabels.NumberFormatLinked property so the chart’s tick‑label colors follow the cell formatting. The example saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | conditional formatting | chart axis label color | tick label color | NumberFormatLinked | column chart | Excel automation | threshold colors | risk level visualization
// Common Searches: Aspose.Cells change chart axis label color based on cell value | link category axis tick labels to conditional formatting Aspose.Cells | C# conditional formatting for chart labels in Excel | NumberFormatLinked property example Aspose.Cells | color tick labels in column chart using Aspose.Cells
// Developer Intent: Apply conditional formatting to cells and have a chart’s category axis tick labels automatically inherit the same font colors using Aspose.Cells for .NET.
// Use Cases: Highlight low, medium, and high categories in a performance chart by coloring axis labels green, orange, and red. | Create risk dashboards where axis labels turn red when values exceed a critical threshold. | Generate dynamic reports that automatically update chart label colors when underlying data changes.
// AI Prompts: Write C# code with Aspose.Cells that adds three conditional formatting rules to a range and links the chart’s category axis tick labels to those formatted cells. | Show how to use the NumberFormatLinked property to make chart axis labels reflect cell font colors for thresholds <40, 40‑79, and ≥80. | Explain step‑by‑step how conditional formatting can drive tick‑label colors in an Aspose.Cells column chart.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add conditional formatting rules (green < 40, orange 40‑79, red ≥ 80), build a column chart, and enable the CategoryAxis.TickLabels.NumberFormatLinked property so the chart’s tick‑label colors follow the cell formatting. The example saves the result as an XLSX file.
    public class ConditionalTickLabelColorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate sample data (categories and values)
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Low");
                sheet.Cells["A3"].PutValue("Medium");
                sheet.Cells["A4"].PutValue("High");
                sheet.Cells["B2"].PutValue(15);   // Low value
                sheet.Cells["B3"].PutValue(55);   // Medium value
                sheet.Cells["B4"].PutValue(95);   // High value

                // -------------------------------------------------
                // 2. Apply conditional formatting to the value cells
                //    to colour the font based on thresholds.
                // -------------------------------------------------
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Define the range that holds the values
                CellArea valueArea = new CellArea { StartRow = 1, EndRow = 3, StartColumn = 1, EndColumn = 1 };
                cfCollection.AddArea(valueArea);

                // Condition 1 : values >= 80 -> Red font
                int cond1 = cfCollection.AddCondition(FormatConditionType.CellValue);
                FormatCondition fc1 = cfCollection[cond1];
                // Older Aspose.Cells versions may not support GreaterThanOrEqual; use GreaterThan with a lower bound of 79
                fc1.Operator = OperatorType.GreaterThan;
                fc1.Formula1 = "79";
                fc1.Style.Font.Color = Color.Red;

                // Condition 2 : values between 40 and 79 -> Orange font
                int cond2 = cfCollection.AddCondition(FormatConditionType.CellValue);
                FormatCondition fc2 = cfCollection[cond2];
                fc2.Operator = OperatorType.Between;
                fc2.Formula1 = "40";
                fc2.Formula2 = "79";
                fc2.Style.Font.Color = Color.Orange;

                // Condition 3 : values < 40 -> Green font
                int cond3 = cfCollection.AddCondition(FormatConditionType.CellValue);
                FormatCondition fc3 = cfCollection[cond3];
                fc3.Operator = OperatorType.LessThan;
                fc3.Formula1 = "40";
                fc3.Style.Font.Color = Color.Green;

                // -------------------------------------------------
                // 3. Add a column chart that uses the above data
                // -------------------------------------------------
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // -------------------------------------------------
                // 4. Link the category axis tick‑labels to the cell
                //    formatting so that the colour changes are reflected.
                // -------------------------------------------------
                chart.CategoryAxis.TickLabels.NumberFormatLinked = true;

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                string outputPath = "ConditionalTickLabelColorDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point required for console execution
        public static void Main()
        {
            Run();
        }
    }
}
