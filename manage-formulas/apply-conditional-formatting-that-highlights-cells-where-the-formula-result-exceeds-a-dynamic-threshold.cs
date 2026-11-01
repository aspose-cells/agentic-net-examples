// Title: How to use Aspose.Cells for .NET to apply conditional formatting that highlights cells exceeding a dynamic threshold value
// Common Searches: Aspose.Cells C# conditional formatting with cell reference threshold | apply expression‑based conditional formatting to a range using Aspose.Cells .NET | highlight cells in an Excel workbook when value is greater than B1 using Aspose.Cells | dynamic threshold conditional formatting example in C# with Aspose.Cells | programmatically add a conditional formatting rule with a formula in Aspose.Cells
// Tags: conditional formatting expression Aspose.Cells .NET | dynamic threshold formatting Excel Aspose | add conditional formatting to cell range Aspose.Cells | CellArea usage for conditional formatting Aspose | style customization for conditional formatting C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The C# sample creates a workbook, fills A2:A10 with incremental numbers, stores a threshold in B1, and adds an expression‑based conditional formatting rule (A2>$B$1) to the range. Cells that exceed the threshold are highlighted with a light‑green background, and the file is saved as ConditionalFormattingDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data in A2:A10
            for (int i = 0; i < 9; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(i * 10); // A2 = 0, A3 = 10, ...
            }

            // Place a dynamic threshold value in B1
            sheet.Cells["B1"].PutValue(30);

            // Define the range to which the conditional formatting will be applied
            string targetRange = "A2:A10";

            // Create a CellArea for the target range (required by the Add method)
            CellArea area = CellArea.CreateCellArea("A2", "A10");

            // Add a new ConditionalFormatting object for the target range
            int cfIndex = sheet.ConditionalFormattings.Add(); // create empty CF

            // Use dynamic to avoid compile‑time dependency on ConditionalFormatting type
            dynamic cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(area); // assign the range

            // Add a condition based on an expression: cell value > $B$1
            // The expression uses the top‑left cell of the range (A2) as a relative reference
            FormatCondition condition = cf.AddCondition(FormatConditionType.Expression);
            condition.Formula1 = "A2>$B$1";

            // Create a style to highlight cells (e.g., light green background)
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.LightGreen;
            highlightStyle.Pattern = BackgroundType.Solid;
            condition.Style = highlightStyle;

            // Save the workbook
            string outputPath = "ConditionalFormattingDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
