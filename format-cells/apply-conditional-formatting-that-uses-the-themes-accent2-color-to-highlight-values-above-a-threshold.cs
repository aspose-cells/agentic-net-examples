// Title: Use Aspose.Cells for .NET to apply conditional formatting with the theme Accent2 color for values greater than a specified threshold
// Common Searches: Aspose.Cells C# set conditional formatting color to theme Accent2 | how to highlight cells greater than 50 using conditional formatting in Aspose.Cells .NET | apply solid background conditional format based on cell value with Aspose.Cells | define conditional formatting range using CellArea in C# Aspose.Cells | save workbook after adding conditional formatting with theme colors using Aspose.Cells
// Tags: theme accent2 conditional formatting Aspose.Cells | cell value greater than threshold formatting .NET | solid fill style conditional format Aspose.Cells | CellArea range definition for conditional formatting | foreground color set with Color.Blue Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells A1:A10 with multiples of 10, defines that range for conditional formatting, adds a rule that formats cells with values greater than 50, applies a solid fill style using a blue foreground color as a substitute for the theme Accent2 color, and saves the file as ConditionalFormatting_Accent2.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                worksheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
            }

            // Define the range for conditional formatting (A1:A10)
            var area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a new ConditionalFormatting object to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            var conditionalFormatting = worksheet.ConditionalFormattings[cfIndex];
            conditionalFormatting.AddArea(area);

            // Threshold value for highlighting
            double threshold = 50;

            // Add a condition: cell value greater than the threshold
            int conditionIndex = conditionalFormatting.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold.ToString(),
                null); // formula2 not required for this operator

            // Retrieve the condition and set its style (using a theme-like color)
            var condition = conditionalFormatting[conditionIndex];
            var style = condition.Style;
            style.ForegroundColor = Color.Blue; // substitute for theme Accent2 color
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Save the workbook
            string outputPath = "ConditionalFormatting_Accent2.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
