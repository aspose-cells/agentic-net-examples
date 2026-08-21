// Title: Offset a range by 3 rows and 5 columns and clear its comments using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data and comments to A1:C3, offsets that range by three rows and five columns, clears all comments in the resulting range, and saves the file.
// Keywords: Aspose.Cells | C# | GetOffset | ClearComments | offset range | remove cell comments | Excel automation | range manipulation
// Common Searches: Aspose.Cells offset range clear comments C# | How to shift a range and delete comments with Aspose.Cells | GetOffset and ClearComments example | Remove comments from moved range Aspose.Cells | C# Aspose.Cells clear comments in offset range
// Developer Intent: Shift an existing cell range by a defined number of rows and columns and delete any comments that appear in the new location.
// Use Cases: Reposition a data block while automatically discarding comments that would otherwise be copied to the new area. | Prepare a report template where comments must be removed from a dynamically calculated section before publishing. | Batch‑process Excel files to move ranges and clean up residual comments without affecting the original worksheet.
// AI Prompts: Generate C# code with Aspose.Cells that offsets a range by N rows and M columns and then clears all comments in the offset range. | Show an example of using Aspose.Cells GetOffset together with ClearComments, including handling when the offset range contains no comments. | Explain how to clear comments from a shifted range safely, ensuring original comments remain untouched.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, adds sample data and comments to A1:C3, offsets that range by three rows and five columns, clears all comments in the resulting range, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add sample data and comments to the original range A1:C3
            cells["A1"].PutValue("Data A1");
            int commentIdx = worksheet.Comments.Add("A1");
            worksheet.Comments[commentIdx].Note = "Comment on A1";

            cells["B2"].PutValue("Data B2");
            commentIdx = worksheet.Comments.Add("B2");
            worksheet.Comments[commentIdx].Note = "Comment on B2";

            cells["C3"].PutValue("Data C3");
            commentIdx = worksheet.Comments.Add("C3");
            worksheet.Comments[commentIdx].Note = "Comment on C3";

            // Define the original range covering A1:C3
            AsposeRange originalRange = cells.CreateRange("A1:C3");

            // Offset the range by 3 rows and 5 columns
            AsposeRange offsetRange = originalRange.GetOffset(3, 5);

            // Clear all comments within the offset range
            offsetRange.ClearComments();

            // Save the workbook
            workbook.Save("OffsetClearComments.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
