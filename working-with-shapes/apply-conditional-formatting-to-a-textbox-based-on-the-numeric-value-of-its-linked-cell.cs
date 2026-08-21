// Title: Aspose.Cells .NET: Conditional Formatting on a TextBox Linked to a Cell
// Description: Demonstrates how to create a workbook, add a TextBox linked to cell B2, apply cell‑level conditional formatting (value > 50 → red, ≤ 50 → light‑green), read the cell value at runtime, and manually sync the TextBox fill and font colors with the formatting before saving the file.
// Keywords: Aspose.Cells conditional formatting textbox | link textbox to cell Aspose.Cells | textbox fill color based on cell value .NET | shape formatting Aspose.Cells | C# Aspose.Cells KPI dashboard | Aspose.Cells manual shape sync | conditional formatting shapes Aspose
// Common Searches: Aspose.Cells change TextBox color by linked cell value | C# conditional formatting on shapes Aspose.Cells | how to sync textbox fill with cell conditional format Aspose | Aspose.Cells .NET example KPI textbox color | apply conditional formatting to a linked TextBox
// Developer Intent: Create a workbook where a TextBox automatically changes its background and font colors according to the numeric value of its linked cell.
// Use Cases: KPI dashboard: a TextBox turns red when a metric exceeds a threshold and green otherwise. | Status reporting: TextBoxes reflect pass/fail conditions by mirroring cell conditional formats. | Automated reports: highlight critical values by synchronizing shape colors with cell rules.
// AI Prompts: Generate C# code with Aspose.Cells that links a TextBox to cell C5 and sets the fill to orange when the value is between 20 and 40. | Show how to loop through multiple TextBoxes and update each fill color based on its linked cell’s conditional formatting in Aspose.Cells .NET. | Explain the steps to synchronize shape formatting with cell conditional formatting in Aspose.Cells, covering different operators and custom colors.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextboxConditionalFormatting
{
    // Demonstrates how to create a workbook, add a TextBox linked to cell B2, apply cell‑level conditional formatting (value > 50 → red, ≤ 50 → light‑green), read the cell value at runtime, and manually sync the TextBox fill and font colors with the formatting before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Add a TextBox shape
            // -------------------------------------------------
            // Parameters: upper left row, upper left column, width (pixels), height (pixels)
            int textboxIndex = sheet.TextBoxes.Add(2, 1, 200, 80);
            TextBox textbox = sheet.TextBoxes[textboxIndex];

            // Set initial text
            textbox.Text = "Value: ";

            // Link the TextBox to a cell (e.g., B2). The textbox will display the cell's value.
            // Using the LinkedCell property (R1C1 style not required, locale true)
            textbox.LinkedCell = "$B$2";

            // -------------------------------------------------
            // 2. Populate the linked cell with a numeric value
            // -------------------------------------------------
            Cell linkedCell = sheet.Cells["B2"];
            linkedCell.PutValue(30); // Change this value to test different formats

            // -------------------------------------------------
            // 3. Add conditional formatting to the linked cell range
            // -------------------------------------------------
            // The conditional formatting will change the cell's background color.
            // Two rules are added: >50 => Red, <=50 => Green.
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

            // Define the range that the formatting applies to (only the linked cell)
            CellArea area = new CellArea
            {
                StartRow = linkedCell.Row,
                EndRow = linkedCell.Row,
                StartColumn = linkedCell.Column,
                EndColumn = linkedCell.Column
            };
            conditions.AddArea(area);

            // Rule 1: Cell value greater than 50 -> Red background
            int rule1 = conditions.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition fc1 = conditions[rule1];
            fc1.Style.BackgroundColor = Color.Red;

            // Rule 2: Cell value less than or equal to 50 -> Green background
            int rule2 = conditions.AddCondition(FormatConditionType.CellValue, OperatorType.LessOrEqual, "50", null);
            FormatCondition fc2 = conditions[rule2];
            fc2.Style.BackgroundColor = Color.LightGreen;

            // -------------------------------------------------
            // 4. Synchronize TextBox appearance with the conditional format result
            // -------------------------------------------------
            // Aspose.Cells does not automatically apply cell conditional formatting to shapes.
            // Therefore we manually inspect the cell's value and adjust the TextBox fill accordingly.
            // This mimics "conditional formatting on a TextBox based on its linked cell".
            double cellValue = linkedCell.DoubleValue;

            if (cellValue > 50)
            {
                // Red fill for values > 50
                textbox.Fill.SolidFill.Color = Color.Red;
                textbox.Font.Color = Color.White;
            }
            else
            {
                // Green fill for values <= 50
                textbox.Fill.SolidFill.Color = Color.LightGreen;
                textbox.Font.Color = Color.Black;
            }

            // Update the displayed text to include the current value
            textbox.Text = $"Value: {cellValue}";

            // -------------------------------------------------
            // 5. Save the workbook (lifecycle: save)
            // -------------------------------------------------
            workbook.Save("TextboxConditionalFormatting.xlsx");
        }
    }
}
