// Title: Apply conditional formatting with Aspose.Cells for .NET to highlight cells whose values exceed a specified threshold
// Common Searches: Aspose.Cells C# add conditional formatting rule for values greater than a limit | how to highlight Excel cells over 100 using Aspose.Cells .NET API | programmatically set background color for cells above a threshold with Aspose.Cells | conditional formatting cell value greater than threshold example in C# | Aspose.Cells conditional formatting for numeric ranges in a worksheet
// Tags: conditional formatting greater-than rule Aspose.Cells | apply style to high-value cells Aspose.Cells .NET | highlight numeric cells exceeding threshold Excel | set cell background color based on value C# | programmatic conditional formatting Excel Aspose

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), defines the range A1:D10 on the first worksheet, adds a conditional formatting rule that marks cells with values greater than 100 using a light coral solid background, and saves the modified file as output.xlsx.
class ConditionalFormattingExample
{
    static void Main()
    {
        try
        {
            // Define the threshold value
            double threshold = 100.0;

            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load an existing workbook if it exists; otherwise create a new one
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure the workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to which the conditional formatting will be applied (A1:D10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 3
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];
            conditions.AddArea(area);

            // Create a condition: cell value greater than the threshold
            int conditionIndex = conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold.ToString(),
                null);
            FormatCondition condition = conditions[conditionIndex];

            // Define the style to apply when the condition is met (light coral background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightCoral;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
