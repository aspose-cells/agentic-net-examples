// Title: Apply expression‑based conditional formatting in Aspose.Cells for .NET to highlight cells exceeding a numeric threshold
// Common Searches: Aspose.Cells C# apply conditional formatting with an expression formula | How to set red background for cells greater than 100 using Aspose.Cells | Create conditional formatting range A1:A10 in a .NET Excel workbook | Using Aspose.Cells to highlight values above a threshold in C# | Add expression condition to conditional formatting collection Aspose.Cells
// Tags: expression conditional formatting Aspose.Cells | conditional formatting range specification .NET | apply style on formula condition Excel C# | threshold based cell highlighting Aspose.Cells | red background style for conditional format

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, fills cells A1‑A10 with incremental values, defines a conditional formatting range, adds an expression condition (A1>100), applies a red solid background style when the condition is true, and saves the file as ConditionalFormattingResult.xlsx.
class ConditionalFormattingExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 20); // Values: 0,20,40,...,180
            }

            // Define the range to which the conditional formatting will be applied (A1:A10)
            CellArea range = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a conditional formatting collection for the specified range
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];
            fcc.AddArea(range);

            // Define a threshold value
            double threshold = 100;

            // Create an expression condition: cell value > threshold
            string formula = $"A1>{threshold}";

            // Add the condition to the collection
            FormatCondition condition = fcc[0];
            condition.Type = FormatConditionType.Expression;
            condition.Formula1 = formula;

            // Define the style to apply when the condition is true (e.g., red background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Determine output file path
            string outputPath = "ConditionalFormattingResult.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
