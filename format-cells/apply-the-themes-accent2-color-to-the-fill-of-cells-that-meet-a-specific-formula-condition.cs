// Title: C# example: Apply theme Accent2 solid fill to cells greater than 5 using Aspose.Cells conditional formatting
// Common Searches: Aspose.Cells C# set Accent2 theme color for conditional formatting rule | How to apply solid fill to cells with values > 5 using Aspose.Cells .NET | Define a cell range for conditional formatting in Aspose.Cells C# | Create a style with theme color and assign to FormatCondition in Aspose.Cells | Save workbook with conditional formatting as .xlsx using Aspose.Cells
// Tags: conditional formatting accent2 fill Aspose.Cells | conditional formatting greater than value Aspose.Cells | CellArea range definition Aspose.Cells | style creation solid fill Aspose.Cells | save workbook xlsx Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook, fills A1:A10 with numbers 1‑10, defines a CellArea covering that range, builds a solid LightBlue style as a stand‑in for the Accent2 theme color, adds a conditional formatting rule that formats cells with values greater than 5, applies the style to the rule, and saves the file as ConditionalFormatting_Accent2.xlsx.
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

            // Populate sample data in A1:A10
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // Values 1..10
            }

            // Define the range for conditional formatting (A1:A10)
            int startRow = 0;      // Row 1 (zero‑based)
            int startColumn = 0;   // Column A
            int endRow = 9;        // Row 10
            int endColumn = 0;     // Column A

            // Create a style with a solid fill (using a standard color)
            Style accent2Style = workbook.CreateStyle();
            accent2Style.ForegroundColor = Color.LightBlue; // Substitute for theme Accent2
            accent2Style.Pattern = BackgroundType.Solid;    // Solid fill

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Set the target range for the conditional formatting using CellArea
            CellArea area = new CellArea
            {
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = endRow,
                EndColumn = endColumn
            };
            cfCollection.AddArea(area);

            // Add a condition: cell value > 5
            int conditionIndex = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "5",
                null);
            FormatCondition condition = cfCollection[conditionIndex];

            // Apply the fill style to cells that meet the condition
            condition.Style = accent2Style;

            // Save the workbook
            string outputPath = "ConditionalFormatting_Accent2.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
