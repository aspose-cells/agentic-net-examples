// Title: Use Aspose.Cells for .NET to apply conditional formatting that highlights cells exceeding a threshold stored in another cell
// AI Prompts: Write C# code with Aspose.Cells that creates a conditional formatting rule to color cells in a range when their value is greater than the value in cell B1. | Generate a snippet that sets a solid light‑coral background style for cells whose computed result surpasses a dynamic threshold defined in a worksheet cell.
// Common Searches: aspnet aspose.cells conditional formatting cell value greater than reference cell B1 | c# apply conditional formatting based on another cell's value using Aspose.Cells | how to set dynamic threshold for conditional formatting in Excel with Aspose.Cells .NET | example of conditional formatting rule with cell reference in Aspose.Cells C# | color cells exceeding threshold stored in B1 using Aspose.Cells API
// Tags: conditional formatting cell value reference Aspose.Cells | apply style based on dynamic threshold .NET Excel | format condition greater than cell Aspose.Cells | solid background color rule Aspose.Cells | Excel workbook conditional formatting C#

using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    // Demonstrates creating a workbook, populating column A, storing a threshold in B1, and adding a conditional formatting rule that colors cells A2:A10 light coral when their values exceed the dynamic threshold.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in column A (A2:A10)
                for (int i = 1; i <= 9; i++) // rows are zero‑based; row 1 = A2
                {
                    sheet.Cells[i, 0].PutValue(10 * i); // 10,20,...,90
                }

                // Set a dynamic threshold value in cell B1
                double threshold = 45.0;
                sheet.Cells["B1"].PutValue(threshold);

                // Define the range for conditional formatting (A2:A10)
                CellArea formatArea = new CellArea
                {
                    StartRow = 1,   // A2
                    StartColumn = 0,
                    EndRow = 9,     // A10
                    EndColumn = 0
                };

                // Get the conditional formatting collection
                var cfCollection = sheet.ConditionalFormattings;

                // Add a new conditional formatting rule
                int cfIndex = cfCollection.Add();
                var cf = cfCollection[cfIndex];

                // Apply the rule to the defined range
                cf.AddArea(formatArea);

                // Add a condition: cell value > $B$1
                int conditionIndex = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "=$B$1",
                    string.Empty);

                // Retrieve the created condition
                FormatCondition condition = cf[conditionIndex];

                // Define the style to apply when the condition is true
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightCoral;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // Save the workbook
                workbook.Save("ConditionalFormattingDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
