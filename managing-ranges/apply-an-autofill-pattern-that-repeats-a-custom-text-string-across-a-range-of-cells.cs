// Title: C# Aspose.Cells AutoFill: Repeat a Custom Text String Across a Range
// Description: Shows how to create a workbook, place a custom string in a source range (A1:A3), and use Range.AutoFill with AutoFillType.Copy to repeat that text across a larger target range (B1:B12) in C#.
// Keywords: Aspose.Cells | AutoFill | C# | .NET | repeat text in Excel | copy range | Excel automation | populate column | custom string | Range.AutoFill
// Common Searches: Aspose.Cells AutoFill copy custom string | C# repeat same text in Excel column using Aspose | How to autofill range with same value Aspose.Cells | AutoFillType.Copy example .NET | Fill column with repeated label Aspose.Cells
// Developer Intent: Copy a small source range that contains a custom string into a larger target range, thereby repeating the text automatically.
// Use Cases: Create a template where a label or instruction must appear in every row of a column. | Populate a report column with a constant status or comment for each record. | Generate a data‑entry sheet that repeats a static instruction without manual typing.
// AI Prompts: Write C# code using Aspose.Cells to autofill the text "Approved" from A1:A2 into C1:C50 with AutoFillType.Copy. | Explain the differences between AutoFillType.Copy, FillSeries, and FillWithoutFormatting when repeating values in Aspose.Cells. | Create a reusable method that accepts a user‑defined string, source range, and target range, then performs AutoFill in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, place a custom string in a source range (A1:A3), and use Range.AutoFill with AutoFillType.Copy to repeat that text across a larger target range (B1:B12) in C#.
    public class AutoFillCustomTextPattern
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the custom text that will be repeated
                string customText = "Sample";

                // Fill a small source range with the custom text (A1:A3)
                for (int i = 0; i < 3; i++)
                {
                    cells[i, 0].PutValue(customText); // Column 0 = "A"
                }

                // Create the source range (A1:A3)
                Aspose.Cells.Range sourceRange = cells.CreateRange("A1:A3");

                // Define the target range where the pattern will be repeated (B1:B12)
                Aspose.Cells.Range targetRange = cells.CreateRange("B1:B12");

                // AutoFill the target range using the Copy type to repeat the source values
                sourceRange.AutoFill(targetRange, AutoFillType.Copy);

                // Save the workbook (output file)
                string outputPath = "AutoFillCustomTextPattern.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}
