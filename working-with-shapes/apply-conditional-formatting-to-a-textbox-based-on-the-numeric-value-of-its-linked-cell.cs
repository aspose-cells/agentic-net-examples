// Title: Apply Conditional Formatting to a TextBox Linked to a Cell with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, insert numeric values, add a TextBox linked to cell A1, define conditional formatting on A1 (green when > B1, red otherwise), force formula calculation, retrieve the evaluated cell style, and apply the same background color to the TextBox before saving the file.
// Keywords: Aspose.Cells | C# conditional formatting | linked TextBox | cell style synchronization | shape fill color | Excel dashboard | .NET workbook example
// Common Searches: Aspose.Cells change TextBox color based on linked cell | conditional formatting for linked shapes Aspose.Cells | C# copy cell background to TextBox Aspose | how to sync TextBox fill with cell style Aspose.Cells | Aspose.Cells conditional formatting example C#
// Developer Intent: Create a TextBox linked to a worksheet cell and automatically update its fill color to match the cell's conditional formatting result.
// Use Cases: KPI indicator that turns green when a value exceeds a target and red when it falls short. | Interactive dashboard where status TextBoxes reflect the conditional colors of their source cells. | Automated report generation that ensures shape colors stay consistent with evaluated cell styles.
// AI Prompts: Generate C# code using Aspose.Cells to link a TextBox to cell C5 and set its background color according to a conditional rule comparing C5 with D5. | Show an Aspose.Cells .NET example that copies the final style of a conditionally formatted cell to a shape’s fill after calling workbook.CalculateFormula. | Explain how to refresh a linked TextBox’s fill color when the underlying cell’s conditional formatting changes in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConditionalTextbox
{
    // Demonstrates how to create a workbook, insert numeric values, add a TextBox linked to cell A1, define conditional formatting on A1 (green when > B1, red otherwise), force formula calculation, retrieve the evaluated cell style, and apply the same background color to the TextBox before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some numeric values in cells that will be linked to the textbox
            // Cell A1 will be linked to the textbox; its value will drive the formatting
            sheet.Cells["A1"].PutValue(75);   // Change this value to test different formats
            sheet.Cells["B1"].PutValue(50);   // Reference value for conditional rule

            // Add a textbox to the worksheet
            // Parameters: upper left row, upper left column, width (pixels), height (pixels)
            int textboxIndex = sheet.TextBoxes.Add(2, 1, 200, 80);
            TextBox textbox = sheet.TextBoxes[textboxIndex];

            // Set the text of the textbox
            textbox.Text = "Current Value: =A1";

            // Link the textbox to cell A1 so that its value reflects the cell's content
            textbox.LinkedCell = "$A$1";

            // -----------------------------------------------------------------
            // Add conditional formatting to the linked cell (A1)
            // The rule: if A1 > B1 then apply a green fill, else apply a red fill
            // -----------------------------------------------------------------
            // Create a new conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range that the conditional formatting will apply to (only A1)
            CellArea area = new CellArea
            {
                StartRow = 0,   // Row 0 = A1
                EndRow = 0,
                StartColumn = 0, // Column 0 = A
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Condition 1: A1 > B1  (green background)
            int conditionIdx1 = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "=B1",   // Formula1: reference cell B1
                null);   // Formula2 not needed for GreaterThan
            FormatCondition condition1 = cfCollection[conditionIdx1];
            condition1.Style.BackgroundColor = Color.LightGreen;

            // Condition 2: A1 <= B1  (red background)
            int conditionIdx2 = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.LessOrEqual,
                "=B1",
                null);
            FormatCondition condition2 = cfCollection[conditionIdx2];
            condition2.Style.BackgroundColor = Color.LightCoral;

            // -----------------------------------------------------------------
            // Optional: synchronize textbox fill color with the cell's background
            // This step reads the evaluated style of A1 and applies it to the textbox.
            // -----------------------------------------------------------------
            // Force calculation to ensure formulas are up‑to‑date
            workbook.CalculateFormula();

            // Retrieve the style applied to cell A1 after conditional formatting
            Style cellStyle = sheet.Cells["A1"].GetStyle();

            // Apply the same background color to the textbox's fill format
            textbox.Fill.SolidFill.Color = cellStyle.BackgroundColor;

            // Save the workbook
            workbook.Save("ConditionalTextboxDemo.xlsx");
        }
    }
}
