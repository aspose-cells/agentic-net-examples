// Title: C# Aspose.Cells: Conditional Formatting with Dynamic Threshold Formula
// Description: Shows how to use Aspose.Cells for .NET to add an expression‑based conditional format that highlights cells in a range when their values exceed a threshold stored in another cell (e.g., B1). The sample creates data, sets the threshold, defines the range, applies the formula =A2>$B$1, styles matching cells, and saves the workbook.
// Keywords: Aspose.Cells | C# conditional formatting | dynamic threshold | expression formula | highlight cells | Excel automation .NET | format condition type expression | cell reference in formula | conditional formatting API | Aspose.Cells tutorial
// Common Searches: Aspose.Cells conditional formatting formula | C# conditional formatting based on another cell | how to set dynamic threshold in Aspose.Cells | apply expression conditional format Aspose.Cells .NET | highlight cells greater than B1 using Aspose.Cells
// Developer Intent: Apply conditional formatting that automatically highlights cells whose values are greater than a threshold defined in a separate cell.
// Use Cases: Flag sales numbers that surpass a target entered in a control cell. | Identify test scores above a passing mark stored in a reference cell. | Mark inventory quantities that exceed a user‑defined reorder limit. | Show budget items that go over a limit set in a dashboard cell.
// AI Prompts: Write Aspose.Cells C# code to apply conditional formatting that colors cells red when they are less than the value in B1. | Create a reusable method that adds a dynamic threshold conditional format to any worksheet range. | Explain how to bind the threshold cell to a named range so the formatting updates when the name changes. | Generate code to apply multiple conditional formats (greater than, less than) using the same threshold cell.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingDynamicThreshold
{
    // Shows how to use Aspose.Cells for .NET to add an expression‑based conditional format that highlights cells in a range when their values exceed a threshold stored in another cell (e.g., B1). The sample creates data, sets the threshold, defines the range, applies the formula =A2>$B$1, styles matching cells, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (A2:A10)
            for (int i = 1; i <= 9; i++) // rows are zero‑based; row 1 = A2
            {
                cells[i, 0].PutValue(i * 10); // 10,20,...,90
            }

            // Define a dynamic threshold in cell B1 (can be a constant or a formula)
            cells[0, 1].PutValue(45); // Threshold = 45

            // Define the range to which the conditional formatting will be applied (A2:A10)
            CellArea range = new CellArea
            {
                StartRow = 1,   // A2
                EndRow = 9,     // A10
                StartColumn = 0,
                EndColumn = 0
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Add an Expression type condition:
            // Formula: =A2>$B$1  (relative to each cell in the range)
            // OperatorType.None is used for Expression type
            int[] result = fcc.Add(
                range,
                FormatConditionType.Expression,
                OperatorType.None,
                "=A2>$B$1",
                null);

            // Retrieve the created FormatCondition (first element of result array)
            FormatCondition fc = fcc[result[0]];

            // Set the style to highlight cells that meet the condition
            fc.Style.BackgroundColor = Color.Yellow;
            fc.Style.Font.Color = Color.Black;
            fc.Style.Font.IsBold = true;

            // Save the workbook
            workbook.Save("DynamicThresholdConditionalFormatting.xlsx");
        }
    }
}
