// Title: Copy a single worksheet row while preserving formulas, comments, and formatting with Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies row 2 to row 5 in an Aspose.Cells worksheet, keeping formulas, comments, and cell styles intact. | Show how to use Aspose.Cells PasteOptions to duplicate an entire row with all attributes preserved. | Provide a step‑by‑step example of copying a source row to a destination row using Range.Copy in C#.
// Common Searches: Aspose.Cells copy row preserving formulas and comments C# | How to duplicate an Excel row with formatting using Aspose.Cells .NET | Copy entire row with PasteType.All in Aspose.Cells example | C# copy worksheet row including cell comments Aspose.Cells | Range.CreateRange copy row to another row Aspose.Cells tutorial
// Tags: copy row with PasteOptions Aspose.Cells | preserve formulas when copying rows .NET | duplicate Excel row including comments C# | range.Copy entire row Aspose.Cells | copy row preserving formatting Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills row 2 with values, a formula, and a comment, then uses a source and destination Range together with PasteOptions (PasteType.All) to copy the whole row to row 5 while retaining formulas, comments, and formatting, and finally saves the file as CopyRowResult.xlsx.
class CopyRowExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source row (row index 1) with data, a formula, and a comment
            cells["A2"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["C2"].Formula = "=A2+B2";

            // Add a comment to a cell in the source row
            Comment comment = sheet.Comments[sheet.Comments.Add("C2")];
            comment.Note = "Sum of A2 and B2";

            // Define source and destination row indices (0‑based)
            int sourceRow = 1; // Excel row 2
            int destRow = 4;   // Excel row 5

            // Determine the total number of columns to copy (covers the whole row)
            int totalColumns = cells.MaxColumn + 1;

            // Create source and destination ranges that represent the entire rows
            Aspose.Cells.Range sourceRange = cells.CreateRange(sourceRow, 0, 1, totalColumns);
            Aspose.Cells.Range destRange   = cells.CreateRange(destRow,   0, 1, totalColumns);

            // Set paste options to copy everything (values, formulas, formats, comments, etc.)
            PasteOptions options = new PasteOptions
            {
                PasteType = PasteType.All
            };

            // Copy the source row to the destination row preserving all attributes
            destRange.Copy(sourceRange, options);

            // Save the workbook
            workbook.Save("CopyRowResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
