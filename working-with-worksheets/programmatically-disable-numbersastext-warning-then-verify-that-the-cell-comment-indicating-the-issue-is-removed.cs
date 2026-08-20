// Title: Disable Numbers Stored As Text warning and delete its comment with Aspose.Cells for .NET
// Description: Demonstrates how to turn off the NumberStoredAsText error check for a specific cell using Aspose.Cells' ErrorCheckOptionCollection, remove the automatically generated warning comment, verify the comment count before and after, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells disable NumberStoredAsText | C# remove Excel warning comment | ErrorCheckOptionCollection example | turn off NumbersAsText error check | programmatic comment deletion Aspose | suppress green triangle warning | Excel numbers stored as text handling
// Common Searches: How to suppress NumberStoredAsText warning in Aspose.Cells | Remove warning comment after disabling NumbersAsText in C# | Aspose.Cells API to turn off NumbersAsText error check for a cell | C# code to delete Excel warning comment with Aspose
// Developer Intent: Programmatically disable the NumberStoredAsText validation for a target cell and erase the associated warning comment.
// Use Cases: Clean up generated workbooks by removing green‑triangle warnings for intentional text‑numeric values. | Prepare Excel files for downstream systems that cannot handle Excel's error‑check indicators. | Automate data import scripts that store numeric identifiers as text without triggering Excel warnings.
// AI Prompts: Generate C# code that disables the NumbersAsText warning for a given range using Aspose.Cells and confirms the comment is removed. | Explain the relationship between ErrorCheckOptionCollection and ErrorCheckOption when turning off NumberStoredAsText validation. | Show how to programmatically check comment count before and after removing a warning comment in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Demonstrates how to turn off the NumberStoredAsText error check for a specific cell using Aspose.Cells' ErrorCheckOptionCollection, remove the automatically generated warning comment, verify the comment count before and after, and save the workbook as an .xlsx file.
class DisableNumbersAsTextWarning
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a number stored as text in cell A1
        sheet.Cells["A1"].PutValue("123"); // value is a string, so Excel would flag it as NumberStoredAsText

        // Add a comment to A1 that represents the warning comment
        int commentIdx = sheet.Comments.Add("A1");
        sheet.Comments[commentIdx].Note = "Number stored as text warning";

        // Verify that the comment exists before disabling the warning
        Console.WriteLine("Comments before disabling warning: " + sheet.Comments.Count); // Expected: 1

        // Disable the NumbersAsText warning for the range that contains A1
        ErrorCheckOptionCollection errorCheckOptions = sheet.ErrorCheckOptions;
        int optionIndex = errorCheckOptions.Add();                     // create a new error‑check option
        ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];
        // Turn off the specific error check type
        errorCheckOption.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false);
        // Apply the option to cell A1
        CellArea cellArea = CellArea.CreateCellArea("A1", "A1");
        errorCheckOption.AddRange(cellArea);

        // After disabling the warning, remove the comment that indicated the issue
        sheet.Comments.RemoveAt("A1");

        // Verify that the comment has been removed
        Console.WriteLine("Comments after disabling warning: " + sheet.Comments.Count); // Expected: 0

        // Save the workbook
        workbook.Save("DisabledNumbersAsText.xlsx");
    }
}
