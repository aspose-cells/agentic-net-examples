// Title: Offset a Range by 3 Rows & 5 Columns and Clear Comments – Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a workbook, add cell comments, define a range (A1:B2), offset it by three rows and five columns using GetOffset, clear all comments inside the shifted range with ClearComments, and save the file.
// Keywords: Aspose.Cells | C# | GetOffset | ClearComments | offset range | remove cell comments | Excel automation | range manipulation | Aspose.Cells for .NET example | Excel comment cleanup
// Common Searches: Aspose.Cells offset range and clear comments C# | How to use GetOffset with ClearComments in Aspose.Cells | Remove comments from a shifted range Aspose.Cells .NET | Aspose.Cells example: offset range then clear comments | C# code to move a range and delete its comments
// Developer Intent: Shift an existing cell range and delete any comments that reside in the new location.
// Use Cases: Reposition a data block in a template while ensuring no leftover comments remain. | Programmatically clean up comments after moving ranges during report generation. | Prepare a worksheet for new data by moving placeholder ranges and clearing their annotations.
// AI Prompts: Write C# code with Aspose.Cells that offsets a range by N rows and M columns and clears all comments in the resulting area. | Show how GetOffset and ClearComments can be combined in Aspose.Cells to move a range and remove its comments, handling empty comment scenarios. | Explain the steps and considerations when using Aspose.Cells to offset a range and clear comments before saving the workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a workbook, add cell comments, define a range (A1:B2), offset it by three rows and five columns using GetOffset, clear all comments inside the shifted range with ClearComments, and save the file.
    public class OffsetClearCommentsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
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

            // Add sample comments to the original range (A1:B2)
            int commentIdx = worksheet.Comments.Add("A1");
            worksheet.Comments[commentIdx].Note = "Comment on A1";

            commentIdx = worksheet.Comments.Add("B2");
            worksheet.Comments[commentIdx].Note = "Comment on B2";

            // Create the original range covering A1:B2
            AsposeRange originalRange = cells.CreateRange("A1:B2");

            // Offset the range by 3 rows and 5 columns
            AsposeRange offsetRange = originalRange.GetOffset(3, 5);

            // Clear all comments within the offset range
            offsetRange.ClearComments();

            // Save the workbook
            workbook.Save("OffsetClearCommentsDemo.xlsx");
            Console.WriteLine("Workbook saved as OffsetClearCommentsDemo.xlsx");
        }
    }
}
