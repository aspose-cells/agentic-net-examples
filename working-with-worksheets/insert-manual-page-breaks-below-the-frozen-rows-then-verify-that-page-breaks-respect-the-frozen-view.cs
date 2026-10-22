// Title: Insert a manual horizontal page break below frozen rows and verify its placement with Aspose.Cells for .NET (C#)
// AI Prompts: Insert a horizontal page break immediately after the frozen rows and output true if the break exists at row index 5. | Freeze the first five rows, add a manual break below them, then iterate the HorizontalPageBreaks collection to confirm the break's row position using Aspose.Cells in C#.
// Common Searches: C# Aspose.Cells add page break after freezing rows | how to confirm page break location after FreezePanes in .NET | sample code for manual page break insertion below frozen rows
// Tags: insert page break Aspose.Cells C# | freeze top rows Aspose.Cells .NET | verify break position Aspose.Cells | manual break insertion after FreezePanes | horizontal break verification C#

using Aspose.Cells;
using System;

// The example creates a new workbook, freezes the top five rows, inserts a horizontal page break right after the frozen area (row index 5), checks the HorizontalPageBreaks collection to ensure the break was added at the expected position, prints the verification result, and saves the file as FrozenRowsWithPageBreak.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Demo";

            // Freeze the top 5 rows (rows 0‑4). The overload requires totalRows and totalColumns parameters.
            sheet.FreezePanes(5, 0, 0, 0);

            // Insert a manual horizontal page break right after the frozen rows (after row 5)
            // Row index is zero‑based, so 5 corresponds to the 6th row in Excel UI.
            sheet.HorizontalPageBreaks.Add(5);

            // Verify that the page break is correctly placed below the frozen rows
            bool breakExists = false;
            foreach (HorizontalPageBreak hpb in sheet.HorizontalPageBreaks)
            {
                if (hpb.Row == 5)
                {
                    breakExists = true;
                    break;
                }
            }

            Console.WriteLine(breakExists
                ? "Page break correctly placed below frozen rows."
                : "Page break not found at expected position.");

            // Save the workbook
            workbook.Save("FrozenRowsWithPageBreak.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
