// Title: How to apply conditional formatting in Aspose.Cells for .NET to set font color to the theme's Dark2 accent for negative values
// Common Searches: Aspose.Cells C# conditional formatting set font color for values less than zero | Change font color to theme Accent2 for negative numbers using Aspose.Cells .NET | Apply conditional formatting to range A1:A20 based on negative values in Aspose.Cells | Create style with theme accent color for conditional formatting in Aspose.Cells C# | How to highlight negative numbers with theme Dark2 font color using Aspose.Cells
// Tags: conditional formatting negative numbers Aspose.Cells | theme accent2 font style Aspose.Cells | apply conditional formatting to range A1:A20 Aspose.Cells | create style with theme color Aspose.Cells .NET | cell value less than zero conditional style .NET

using System;
using System.IO;
using System.Drawing;               // For Color
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, defines a conditional formatting rule for cells A1:A20 that triggers when the cell value is less than zero, and applies a style that sets the font color to an approximate theme Dark2 (Accent2) color before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Define the range for conditional formatting (A1:A20)
            string startCell = "A1";
            string endCell = "A20";

            // Add a new conditional formatting rule to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            var cf = worksheet.ConditionalFormattings[cfIndex];

            // Set the area the rule applies to
            cf.AddArea(CellArea.CreateCellArea(startCell, endCell));

            // Add a condition: cell value less than 0 (negative numbers)
            int conditionIndex = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.LessThan,
                "0",
                null);

            // Retrieve the created condition object
            var condition = cf[conditionIndex];

            // Create a style that uses a color similar to the theme's Accent2 for the font
            var style = workbook.CreateStyle();
            style.Font.Color = Color.FromArgb(0, 112, 192); // Approximate Accent2 color

            // Apply the style to the condition
            condition.Style = style;

            // Define output file path
            string outputPath = "ConditionalFormatting_Negative_Accent2.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
