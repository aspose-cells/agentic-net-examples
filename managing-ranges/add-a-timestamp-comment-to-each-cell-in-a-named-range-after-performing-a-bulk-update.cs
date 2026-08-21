// Title: Add timestamp comments to each cell in a named range after a bulk update – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a named range (A1:C3), bulk‑updates all cells with a single value, then adds a comment containing the current date‑time to every cell in the range, and saves the file as TimestampComments.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timestamp comment | C# add comment to range | named range bulk update Aspose.Cells | Excel comment with date time | programmatic comment each cell | Aspose.Cells .NET example | Excel automation timestamp
// Common Searches: Aspose.Cells add comment with current date to each cell | bulk update named range and insert timestamp comments C# | how to add timestamp comment to a range using Aspose.Cells | C# Aspose.Cells comment per cell after PutValue | save Excel with timestamped comments Aspose
// Developer Intent: Apply a single value to all cells in a named range and automatically attach a date‑time comment to each cell for audit or tracking purposes.
// Use Cases: Record the exact moment data was refreshed in a financial report. | Provide per‑cell audit trails in spreadsheets that undergo batch processing. | Add change‑history notes to exported Excel files for downstream consumers.
// AI Prompts: Generate C# code with Aspose.Cells that updates a named range in bulk and adds a timestamp comment to every cell. | Show how to retrieve a named range, perform PutValue on the whole range, then iterate cells to set a comment with DateTime.Now. | Explain best practices for formatting timestamp comments and ensuring they persist after saving the workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTimestampCommentDemo
{
    // Creates a workbook, defines a named range (A1:C3), bulk‑updates all cells with a single value, then adds a comment containing the current date‑time to every cell in the range, and saves the file as TimestampComments.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 1. Define a named range (for demonstration purposes)
                // -----------------------------------------------------------------
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // The named range refers to cells A1:C3 on the first sheet
                namedRange.RefersTo = "=Sheet1!$A$1:$C$3";

                // -----------------------------------------------------------------
                // 2. Retrieve the range object from the named range definition
                // -----------------------------------------------------------------
                // RefersTo includes the leading '=', remove it to obtain the address
                string address = namedRange.RefersTo.TrimStart('=');
                // Create a Range object based on the address
                AsposeRange range = worksheet.Cells.CreateRange(address);

                // -----------------------------------------------------------------
                // 3. Perform a bulk update on the range (e.g., set the same value)
                // -----------------------------------------------------------------
                // PutValue requires conversion and style flags; set both to false for plain text
                range.PutValue("Bulk Updated", false, false);

                // -----------------------------------------------------------------
                // 4. Add a timestamp comment to each cell in the range
                // -----------------------------------------------------------------
                foreach (Cell cell in range)
                {
                    int row = cell.Row;
                    int column = cell.Column;

                    // Add a comment to the current cell using the (row, column) overload
                    int commentIndex = worksheet.Comments.Add(row, column);
                    Comment comment = worksheet.Comments[commentIndex];

                    // Set the comment text to include the current timestamp
                    comment.Note = $"Updated on {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                }

                // -----------------------------------------------------------------
                // 5. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "TimestampComments.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
