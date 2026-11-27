// Title: How to programmatically suppress the NumbersAsText warning and delete its comment using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET disable NumbersAsText warning and delete related comment | remove "Number stored as text" comment from a worksheet with Aspose.Cells C# | clear all comments in an Excel workbook using Aspose.Cells API | suppress NumbersAsText warning without affecting other validations in Aspose.Cells | C# code to turn off NumbersAsText warning and verify comment removal
// Tags: NumbersAsText warning suppression Aspose.Cells | worksheet comment removal Aspose.Cells C# | clear Excel comments via Aspose.Cells API | disable number-as-text validation Aspose.Cells | Aspose.Cells comment handling .NET

using System;
using Aspose.Cells;

// The example creates a new Workbook, writes a numeric value as text in cell A1, adds a comment indicating the 'Number stored as text' issue, then clears all worksheet comments to simulate disabling the NumbersAsText warning, verifies that the comment count is zero, and saves the workbook as NumbersAsTextDisabled.xlsx.
class DisableNumbersAsTextWarning
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value as text into cell A1 (this would normally raise a warning)
            sheet.Cells["A1"].PutValue("123");

            // Add a comment indicating the issue
            int commentIndex = sheet.Comments.Add("A1");
            sheet.Comments[commentIndex].Note = "Number stored as text";

            // Verify that the comment was added
            Console.WriteLine("Comments before disabling warning: " + sheet.Comments.Count); // Expected: 1

            // Simulate disabling the NumbersAsText warning by clearing comments
            sheet.Comments.Clear();

            // Verify that comments have been removed
            bool commentRemoved = sheet.Comments.Count == 0;
            Console.WriteLine("Comment removed after disabling warning: " + commentRemoved); // Expected: True

            // Save the workbook
            string outputPath = "NumbersAsTextDisabled.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
