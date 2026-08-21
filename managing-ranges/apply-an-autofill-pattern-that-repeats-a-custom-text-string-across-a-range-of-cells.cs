// Title: Repeat a custom text string across a cell range with Aspose.Cells AutoFill in C#
// Description: Creates a workbook, writes a custom string to A1, defines A1 as the source range and A2:A11 as the target range, then uses AutoFill with AutoFillType.Copy to duplicate the text across the target cells and saves the file.
// Keywords: Aspose.Cells | AutoFill | AutoFillType.Copy | C# | .NET | repeat text across cells | fill range with custom string | Excel automation | workbook | cells range
// Common Searches: Aspose.Cells repeat same text in a column | AutoFill copy single cell value to range C# | How to fill A2:A10 with text from A1 using Aspose.Cells | C# Aspose.Cells autofill custom string vertically | Copy cell value to multiple cells Aspose.Cells .NET
// Developer Intent: Copy a single cell's text value into a larger range using Aspose.Cells AutoFill.
// Use Cases: Populate a label column in a generated report where every row needs the same heading. | Create a template that automatically inserts a warning or instruction text into each data entry row. | Initialize a worksheet with a repeated note for user guidance across many rows.
// AI Prompts: Show how to modify the code to autofill the custom text horizontally across cells B1:G1. | Provide an example that uses AutoFillType.FillSeries to add sequential numbers after the custom text in a column. | Explain how to set up an AutoFill pattern that alternates two different strings across a range of cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook, writes a custom string to A1, defines A1 as the source range and A2:A11 as the target range, then uses AutoFill with AutoFillType.Copy to duplicate the text across the target cells and saves the file.
    public class AutoFillCustomTextDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the custom text to repeat
            string customText = "Sample Text";

            // Put the custom text into a single source cell (A1)
            cells["A1"].PutValue(customText);

            // Create the source range that contains the text (A1)
            AsposeRange sourceRange = cells.CreateRange("A1");

            // Define the target range where the text should be repeated (A2:A11)
            AsposeRange targetRange = cells.CreateRange("A2:A11");

            // Use AutoFill with the Copy type to repeat the text across the target range
            sourceRange.AutoFill(targetRange, AutoFillType.Copy);

            // Save the workbook to a file
            workbook.Save("AutoFillCustomTextDemo.xlsx");
        }
    }
}
