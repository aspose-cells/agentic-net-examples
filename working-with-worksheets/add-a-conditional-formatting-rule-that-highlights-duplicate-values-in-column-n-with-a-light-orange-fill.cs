// Title: C# Aspose.Cells example: add conditional formatting to highlight duplicate values in column N with a light orange fill
// Common Searches: Aspose.Cells C# highlight duplicate entries in column N | How to apply orange fill to duplicate cells using Aspose.Cells .NET | Create conditional formatting for duplicate values in Excel with Aspose.Cells C# | Add duplicate value conditional formatting to a specific column in C# Aspose.Cells | Set solid orange background for duplicate values in Excel via Aspose.Cells API
// Tags: Aspose.Cells duplicate values conditional formatting | C# conditional formatting orange fill Excel | highlight duplicates column N Aspose.Cells | set solid background color Aspose.Cells style | Excel conditional formatting API .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a workbook (if missing), fills column N with duplicate values, then uses Aspose.Cells for .NET to add a conditional formatting rule that highlights those duplicates with a light orange solid fill, and saves the result as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a sample workbook if it does not.
            if (!File.Exists(inputPath))
            {
                try
                {
                    var sampleWorkbook = new Workbook();
                    var sheet = sampleWorkbook.Worksheets[0];
                    sheet.Name = "Sheet1";

                    // Populate column N (index 13) with some duplicate values for demonstration.
                    for (int row = 0; row < 10; row++)
                    {
                        sheet.Cells[row, 13].PutValue(row % 3);
                    }

                    sampleWorkbook.Save(inputPath);
                    Console.WriteLine($"Sample input file created at '{inputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample input file: {ex.Message}");
                    return;
                }
            }

            try
            {
                // Load the workbook.
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // Add a new conditional formatting rule.
                int cfIndex = worksheet.ConditionalFormattings.Add();
                var conditionalFormatting = worksheet.ConditionalFormattings[cfIndex];

                // Define the range for column N (entire column).
                int lastRow = worksheet.Cells.MaxDataRow;
                var area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 13,
                    EndRow = lastRow,
                    EndColumn = 13
                };
                conditionalFormatting.AddArea(area);

                // Create a condition that highlights duplicate values.
                int conditionIndex = conditionalFormatting.AddCondition(FormatConditionType.DuplicateValues);

                // Set the style for the condition (light orange fill).
                Style style = conditionalFormatting[conditionIndex].Style;
                style.ForegroundColor = Color.FromArgb(255, 255, 230, 153);
                style.Pattern = BackgroundType.Solid;

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Conditional formatting applied and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
