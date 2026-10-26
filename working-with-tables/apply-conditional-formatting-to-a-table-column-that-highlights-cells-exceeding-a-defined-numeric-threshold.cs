// Title: Apply conditional formatting to a worksheet column to highlight cells greater than a numeric threshold with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add conditional formatting to column based on value | highlight cells greater than 50 in Excel using Aspose.Cells .NET | set red background for cells exceeding threshold with Aspose.Cells C# | create conditional formatting rule for numeric values in a specific column using Aspose.Cells | conditional formatting B1:B10 Aspose.Cells example
// Tags: apply conditional format to worksheet column Aspose.Cells | greater than numeric threshold condition Aspose.Cells | solid red fill style for conditional format .NET | format condition type CellValue operator GreaterThan | excel .xlsx conditional formatting using C# Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, fills column B (B1:B10) with values 0‑90, defines a threshold of 50, and adds a conditional formatting rule that applies a solid red background to any cell in that column whose value exceeds the threshold, then saves the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column B (index 1)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 1].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Define the numeric threshold
            double threshold = 50;

            // Define the range for the conditional formatting (B1:B10)
            int startRow = 0;
            int endRow = 9;
            int columnIndex = 1; // Column B

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var conditionalFormatting = sheet.ConditionalFormattings[cfIndex];

            // Set the area that the formatting will apply to
            CellArea area = new CellArea
            {
                StartRow = startRow,
                StartColumn = columnIndex,
                EndRow = endRow,
                EndColumn = columnIndex
            };
            conditionalFormatting.AddArea(area);

            // Create a condition: cell value greater than the threshold
            // The AddCondition method requires two formula strings; the second can be empty for this case
            int conditionIndex = conditionalFormatting.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold.ToString(),
                string.Empty);

            // Retrieve the created condition
            FormatCondition condition = conditionalFormatting[conditionIndex];

            // Define the style to apply when the condition is met (red background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Prepare output path
            string outputFile = "ConditionalFormattingExample.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile)) ?? Directory.GetCurrentDirectory();

            // Ensure the output directory exists
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
