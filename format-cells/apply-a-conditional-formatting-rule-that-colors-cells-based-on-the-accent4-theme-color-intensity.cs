// Title: Apply Accent4 theme color with conditional formatting to a cell range in Aspose.Cells for .NET (C#)
// Common Searches: how to use Aspose.Cells C# to apply Accent4 theme color in conditional formatting | Aspose.Cells conditional formatting rule for values greater than 50 with solid fill | C# example of setting conditional formatting background color using theme colors in Aspose.Cells | apply solid LightBlue Accent4 style to range A1:D10 with Aspose.Cells | Aspose.Cells .NET conditional formatting based on cell value threshold
// Tags: Aspose.Cells conditional formatting with theme colors | C# set solid fill style in Aspose.Cells | conditional format cell value greater than 50 Aspose | define conditional formatting range A1:D10 Aspose.Cells | apply Accent4 background style Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, defines the A1:D10 range, adds a conditional formatting rule that triggers when cell values exceed 50, applies a solid LightBlue (Accent4 theme) background style, and saves the file as ConditionalFormatting_Accent4.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range for conditional formatting (A1:D10)
            CellArea area = new CellArea
            {
                StartRow = 0,    // Row 1 (zero‑based)
                EndRow = 9,      // Row 10
                StartColumn = 0, // Column A
                EndColumn = 3    // Column D
            };

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(area);

            // Add a condition: cells with values greater than 50
            int conditionIndex = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                null);

            // Retrieve the created condition
            FormatCondition condition = cf[conditionIndex];

            // Create a style for the conditional formatting
            Style accentStyle = workbook.CreateStyle();
            accentStyle.Pattern = BackgroundType.Solid;
            accentStyle.ForegroundColor = Color.LightBlue;

            // Assign the style to the condition
            condition.Style = accentStyle;

            // Save the workbook
            string outputPath = "ConditionalFormatting_Accent4.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
