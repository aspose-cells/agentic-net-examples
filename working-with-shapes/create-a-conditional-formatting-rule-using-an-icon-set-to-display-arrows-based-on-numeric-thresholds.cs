// Title: How to add a three‑arrow icon set conditional formatting rule to a cell range using Aspose.Cells for C#
// Common Searches: c# apply three arrow icon set conditional formatting with Aspose.Cells | example of using FormatConditionType.IconSet in Aspose.Cells .NET | programmatically set Excel icon set formatting for numeric thresholds in C# | aspnet add icon set conditional formatting to column A using Aspose.Cells | how to create arrow icons conditional formatting for a range in .NET Excel
// Tags: Aspose.Cells icon set formatting API | C# Excel conditional formatting with icon set | Apply three‑arrow icon set to cell range | FormatConditionType.IconSet example | Conditional formatting range definition Aspose

using System;
using Aspose.Cells;

// The sample creates a workbook, writes numeric values to cells A1‑A4, defines that range, adds a FormatConditionType.IconSet conditional formatting rule (default three‑arrow set) to the range, and saves the file as ConditionalIconSet.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some numeric values in column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(30);
            sheet.Cells["A3"].PutValue(70);
            sheet.Cells["A4"].PutValue(90);

            // Define the range (A1:A4) that will receive the icon set conditional formatting
            CellArea range = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A (zero‑based)
                EndRow = 3,        // Row 4
                EndColumn = 0      // Column A
            };

            // Add a new ConditionalFormatting object to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();

            // Associate the defined range with the ConditionalFormatting object
            sheet.ConditionalFormattings[cfIndex].AddArea(range);

            // Add an IconSet condition (no operator needed for icon sets)
            sheet.ConditionalFormattings[cfIndex].AddCondition(FormatConditionType.IconSet, OperatorType.None, "", "");

            // Retrieve the condition we just added
            FormatCondition iconCondition = sheet.ConditionalFormattings[cfIndex][0];

            // The default icon set for IconSet type is a three‑arrow set, so no explicit assignment is required.
            // If a different set is needed, it can be set via the IconSet property (if supported by the library version).

            // Save the workbook to a file
            string outputPath = "ConditionalIconSet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
