using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.mht";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell area for conditional formatting (A1:B10)
            CellArea formatArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 1
            };

            // Add a new conditional formatting collection
            int cfIndex = worksheet.ConditionalFormattings.Add();
            var conditionalFormatting = worksheet.ConditionalFormattings[cfIndex];

            // Associate the area with the conditional formatting
            conditionalFormatting.AddArea(formatArea);

            // Add a condition: cell value greater than 50
            int conditionIndex = conditionalFormatting.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                string.Empty); // formula2 not required for this operator

            // Create a style (yellow background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the condition
            conditionalFormatting[conditionIndex].Style = style;

            // Save as MHTML (MHT) web archive
            workbook.Save(outputPath, SaveFormat.MHtml);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}