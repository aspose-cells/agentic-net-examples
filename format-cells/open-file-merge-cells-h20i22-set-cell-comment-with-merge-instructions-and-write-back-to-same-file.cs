// Title: Merge cells H20:I22, add a comment, and overwrite the original Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a specific range and add a comment before saving | overwrite existing Excel workbook after merging cells with Aspose.Cells .NET | add a note to merged cells H20 I22 using Aspose.Cells in C# | load or create workbook then merge H20:I22 and save with Aspose.Cells | C# code to merge cells and attach comment in the same Excel file using Aspose.Cells
// Tags: cell range merging Aspose.Cells C# | attach comment to merged range Aspose.Cells | overwrite original workbook Aspose.Cells | H20:I22 range handling Aspose.Cells | load or create workbook Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new one if it doesn't exist), merges the range H20:I22 on the first worksheet, adds a comment to the top‑left cell of the merged area, and then saves the workbook back to the same file, overwriting the original.
class Program
{
    static void Main()
    {
        // Path to the Excel file (will be created if it does not exist)
        string filePath = "input.xlsx";

        Workbook workbook = null;

        // Load existing workbook if the file exists; otherwise create a new one
        if (File.Exists(filePath))
        {
            try
            {
                workbook = new Workbook(filePath);
            }
            catch (CellsException ex)
            {
                // Handle cases such as password‑protected files
                Console.WriteLine($"Unable to load workbook \"{filePath}\": {ex.Message}");
                Console.WriteLine("A new workbook will be created instead.");
                workbook = new Workbook();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error loading workbook: {ex.Message}");
                workbook = new Workbook();
            }
        }
        else
        {
            workbook = new Workbook(); // create a new workbook with a default worksheet
        }

        try
        {
            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells H20:I22 (zero‑based column 7, row 19, 3 rows, 2 columns)
            sheet.Cells.Merge(7, 19, 3, 2);

            // Add a comment to the top‑left cell of the merged range (H20)
            int commentIndex = sheet.Comments.Add("H20");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Merged cells H20:I22";

            // Save the workbook (creates the file if it did not exist)
            workbook.Save(filePath);
            Console.WriteLine($"Workbook saved successfully to \"{filePath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
