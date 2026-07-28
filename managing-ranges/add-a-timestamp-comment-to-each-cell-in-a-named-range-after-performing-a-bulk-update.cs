// Title: Add Timestamp Comments to All Cells in a Named Range After Bulk Update (C# Aspose.Cells)
// Description: Creates a workbook, defines a named range, performs a bulk cell update, parses the range's RefersTo string, and adds a comment with the current date‑time to each cell in the range before saving the file.
// Keywords: Aspose.Cells | C# | .NET | timestamp comment | named range | bulk update | add comment to cells | Excel comment programmatically | RefersTo parsing | audit trail
// Common Searches: Aspose.Cells add comment to range C# | C# timestamp comment named range Aspose.Cells | How to bulk update cells and insert comments with Aspose.Cells | Parse named range RefersTo Aspose.Cells .NET | Add Excel comments to multiple cells using Aspose.Cells
// Developer Intent: Insert a current date‑time comment into every cell of a specified named range after a bulk data change using Aspose.Cells for .NET.
// Use Cases: Generate an audit log of when each cell was modified during a data transformation. | Tag imported bulk data with processing timestamps for downstream validation. | Show end‑users the exact time a cell was updated in automatically generated reports. | Enable change tracking in shared Excel files by recording edit timestamps.
// AI Prompts: Write C# code with Aspose.Cells that adds a comment containing the current timestamp to each cell in a named range. | Explain how to extract the worksheet name and address from a named range's RefersTo string and use it to create a Range object. | Show how to handle existing comments when adding a new timestamp comment to cells in Aspose.Cells. | Provide a sample that formats the timestamp as "yyyy-MM-dd HH:mm:ss" and inserts it into cell comments. | Demonstrate how to combine a bulk update (e.g., value multiplication) with comment insertion in a single Aspose.Cells workflow.

using System;
using Aspose.Cells;

namespace AsposeCellsTimestampCommentDemo
{
    // Creates a workbook, defines a named range, performs a bulk cell update, parses the range's RefersTo string, and adds a comment with the current date‑time to each cell in the range before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data in the worksheet (A1:C3)
                for (int r = 0; r < 3; r++)
                {
                    for (int c = 0; c < 3; c++)
                    {
                        sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                    }
                }

                // Define a named range that covers A1:C3
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo must include the sheet name and absolute references
                namedRange.RefersTo = $"={sheet.Name}!$A$1:$C$3";

                // ----- Bulk update -----
                // Example bulk operation: multiply each numeric cell by 10
                // (Here we just set a new value for demonstration)
                Aspose.Cells.Range bulkRange = sheet.Cells.CreateRange("A1", "C3");
                foreach (Cell cell in bulkRange)
                {
                    // Simple bulk update: append "-Updated" to existing text
                    cell.PutValue($"{cell.StringValue}-Updated");
                }

                // ----- Add timestamp comment to each cell in the named range -----
                // Resolve the range object from the named range's RefersTo string
                // RefersTo format: =SheetName!$A$1:$C$3
                string refersTo = namedRange.RefersTo.TrimStart('=');
                int exclPos = refersTo.IndexOf('!');
                string sheetName = refersTo.Substring(0, exclPos);
                string address = refersTo.Substring(exclPos + 1);

                // Get the worksheet that the named range refers to
                Worksheet targetSheet = workbook.Worksheets[sheetName];

                // Create a Range object for the address
                Aspose.Cells.Range targetRange = targetSheet.Cells.CreateRange(address);

                // Access the comment collection of the target worksheet
                CommentCollection comments = targetSheet.Comments;

                // Current timestamp string
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Iterate through each cell in the range and add a comment
                foreach (Cell cell in targetRange)
                {
                    int row = cell.Row;
                    int col = cell.Column;

                    // Add a comment (if a comment already exists, this will create a new one)
                    int commentIndex = comments.Add(row, col);
                    Comment comment = comments[commentIndex];
                    comment.Note = $"Updated on {timestamp}";
                }

                // Save the workbook
                workbook.Save("TimestampCommentDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
