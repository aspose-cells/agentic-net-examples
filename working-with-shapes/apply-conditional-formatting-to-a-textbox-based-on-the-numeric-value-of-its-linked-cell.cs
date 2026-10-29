// Title: Apply conditional formatting to a TextBox shape based on its linked cell value using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set TextBox font color based on linked cell value | Conditional formatting for a TextBox shape linked to a worksheet cell in Aspose.Cells .NET | Change TextBox appearance using cell value conditions with Aspose.Cells | Create a conditional formatting rule for a linked TextBox in Excel using Aspose.Cells C# | How to color a TextBox in Excel according to its linked cell using Aspose.Cells
// Tags: Aspose.Cells shape conditional formatting | C# change TextBox font color by cell value | Excel conditional rule applied to shape linked cell | Apply style to cell area using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a TextBox shape linked to cell A1, defines two conditional formatting rules for that cell (red font when the value is greater than 10, green font when less than or equal to 10), and saves the result as 'ConditionalTextbox.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some numeric values in cells
            sheet.Cells["A1"].PutValue(5);   // Value that will trigger the "less or equal" condition
            sheet.Cells["A2"].PutValue(15);  // Another cell for demonstration (not linked)

            // Add a TextBox shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, height, width
            TextBox textBox = sheet.Shapes.AddTextBox(5, 0, 0, 0, 200, 50);

            // Link the TextBox to cell A1 so it displays the cell's value
            textBox.LinkedCell = "A1";

            // Apply conditional formatting to the linked cell (A1)
            // Create a conditional formatting rule for the range A1
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the cell area the rule applies to (A1)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 0,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Condition 1: If value > 10, set font color to Red
            int redCondIdx = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "10",
                null);
            FormatCondition conditionRed = cfCollection[redCondIdx];
            Style styleRed = workbook.CreateStyle();
            styleRed.Font.Color = Color.Red;
            conditionRed.Style = styleRed;

            // Condition 2: If value <= 10, set font color to Green
            int greenCondIdx = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.LessOrEqual,
                "10",
                null);
            FormatCondition conditionGreen = cfCollection[greenCondIdx];
            Style styleGreen = workbook.CreateStyle();
            styleGreen.Font.Color = Color.Green;
            conditionGreen.Style = styleGreen;

            // Determine output file path and ensure directory exists
            string outputPath = "ConditionalTextbox.xlsx";
            string fullPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
