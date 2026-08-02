// Title: Aspose.Cells for .NET – Disable NumbersAsText (NumberStoredAsText) Warning and Remove Its Comment
// Description: Demonstrates how to turn off the NumbersAsText (NumberStoredAsText) error check for a specific cell range using Aspose.Cells, verify that the automatically added comment disappears, and save the workbook.
// Keywords: Aspose.Cells | C# | disable NumbersAsText warning | NumberStoredAsText | ErrorCheckOptionCollection | remove cell comment | worksheet error check | suppress numeric-as-text error | Excel warning suppression
// Common Searches: how to disable NumbersAsText warning in Aspose.Cells | remove comment created by NumberStoredAsText error check | Aspose.Cells turn off numeric stored as text for a range | verify comment count after disabling error check Aspose.Cells | C# Aspose.Cells suppress NumbersAsText error
// Developer Intent: Turn off the NumbersAsText (NumberStoredAsText) error check for a given range and confirm that the generated comment is removed.
// Use Cases: Import data that intentionally stores numbers as text without triggering warnings. | Clean up worksheet comments after adjusting error‑check settings programmatically. | Apply distinct error‑check configurations to multiple ranges within the same sheet.
// AI Prompts: Write C# code with Aspose.Cells to disable the NumberStoredAsText warning for cells B2:B10 and ensure no comments remain. | Explain the role of ErrorCheckOptionCollection in Aspose.Cells and how to assign multiple ranges to a disabled warning. | Provide a step‑by‑step guide to check the comment count before and after turning off a warning in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to turn off the NumbersAsText (NumberStoredAsText) error check for a specific cell range using Aspose.Cells, verify that the automatically added comment disappears, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a numeric value as text in cell A1 (using a leading apostrophe)
        worksheet.Cells["A1"].PutValue("'123");

        // At this point the default error check adds a comment for NumbersAsText
        Console.WriteLine($"Comments before disabling warning: {worksheet.Comments.Count}");

        // Disable the NumbersAsText (NumberStoredAsText) warning for the range A1
        ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;
        int optionIndex = errorCheckOptions.Add();                     // create a new option
        ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];
        errorCheckOption.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false); // turn off the warning
        CellArea cellArea = CellArea.CreateCellArea("A1", "A1");       // define the range to which the option applies
        errorCheckOption.AddRange(cellArea);

        // After disabling the warning the comment should be removed automatically
        Console.WriteLine($"Comments after disabling warning: {worksheet.Comments.Count}");

        // Save the workbook (lifecycle rule)
        workbook.Save("NumbersAsTextDisabled.xlsx");
    }
}
