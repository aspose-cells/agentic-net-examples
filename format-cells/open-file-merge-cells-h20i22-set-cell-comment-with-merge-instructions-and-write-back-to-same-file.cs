// Title: Merge cells H20:I22 and add a comment in an existing Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge range H20 to I22 and insert a comment before saving | how to programmatically merge specific cells and add a note in an existing .xlsx with Aspose.Cells | C# example for adding a comment to the top‑left cell of a merged range in Aspose.Cells | overwrite an existing Excel file after merging cells and setting a comment using Aspose.Cells .NET | add a cell comment to a merged area in Aspose.Cells workbook via C# code
// Tags: cell range merging Aspose.Cells | insert comment into merged cells C# | save changes to existing .xlsx Aspose.Cells | programmatic Excel modification .NET | Aspose.Cells workbook cell comment

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample loads an existing Excel workbook (or creates a new one), merges the cells H20:I22 on the first worksheet, adds a comment to the top‑left cell of the merged area, and saves the workbook back to the same file, overwriting the original.
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "YourFile.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(filePath))
                {
                    workbook = new Workbook(filePath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Merge cells H20:I22 (zero‑based indices: row 19, column 7, 3 rows, 2 columns)
                sheet.Cells.Merge(19, 7, 3, 2);

                // Add a comment to the top‑left cell of the merged range (H20)
                int commentIdx = sheet.Comments.Add("H20");
                Comment comment = sheet.Comments[commentIdx];
                comment.Note = "Cells H20:I22 have been merged as per instructions.";

                // Save the workbook (overwrites existing file or creates a new one)
                workbook.Save(filePath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(filePath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
