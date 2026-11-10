// Title: How to use Aspose.Cells for .NET to apply conditional formatting that highlights cells with values greater than a threshold
// Common Searches: Aspose.Cells C# add conditional formatting for cells greater than 50 | C# conditional formatting yellow fill for values above a limit using Aspose.Cells | Apply greater‑than condition to a cell range in an Aspose.Cells workbook | Set up a FormatConditionType.CellValue rule in .NET with Aspose.Cells | Highlight numeric cells exceeding a threshold in an Excel file with Aspose.Cells
// Tags: Aspose.Cells conditional formatting greater than | C# conditional formatting cell value threshold | Aspose.Cells FormatConditionType CellValue | Aspose.Cells solid yellow fill style | Aspose.Cells conditional formatting CellArea range

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills column A with values 0‑90, defines a conditional formatting rule on range A1:A10 that applies a solid yellow fill to cells whose value exceeds 50, and saves the file as ConditionalFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate example data (optional)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Threshold for highlighting
            double threshold = 50.0;

            // Define the range for conditional formatting (A1:A10)
            int startRow = 0;      // zero‑based row index
            int startColumn = 0;   // column A
            int endRow = 9;        // row 10
            int endColumn = 0;     // column A

            // Add a new ConditionalFormatting entry to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();

            // Add the cell range to the conditional formatting
            sheet.ConditionalFormattings[cfIndex].AddArea(new CellArea
            {
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = endRow,
                EndColumn = endColumn
            });

            // Add a condition: cell value greater than the threshold
            // The AddCondition method requires two formula parameters; the second can be empty for this case
            int conditionIndex = sheet.ConditionalFormattings[cfIndex].AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold.ToString(),
                string.Empty);

            // Define the style to apply when the condition is met (yellow fill)
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;

            // Assign the style to the condition
            sheet.ConditionalFormattings[cfIndex][conditionIndex].Style = highlightStyle;

            // Define output file path
            string outputPath = "ConditionalFormatting.xlsx";

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
