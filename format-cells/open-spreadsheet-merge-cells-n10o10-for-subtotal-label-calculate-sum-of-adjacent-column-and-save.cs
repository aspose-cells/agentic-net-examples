// Title: C# Aspose.Cells: Merge N10:O10, add “Subtotal” label, and sum N1:N9 into P10
// Description: Demonstrates loading an Excel file with Aspose.Cells for .NET, merging cells N10:O10, inserting the text “Subtotal”, assigning a SUM(N1:N9) formula to P10, and saving the updated workbook.
// Keywords: Aspose.Cells | C# | .NET | merge cells | Excel subtotal | SUM formula | column total | save workbook | cell merging Aspose | set formula Aspose
// Common Searches: Aspose.Cells merge cells N10 O10 C# | Set SUM formula in P10 using Aspose.Cells | Add Subtotal label to merged cells Aspose.Cells | C# calculate column total with Aspose.Cells | How to save workbook after merging cells Aspose
// Developer Intent: Combine cell merging, label insertion, formula assignment, and workbook saving in a single Aspose.Cells workflow.
// Use Cases: Financial statements where the subtotal row header spans two columns and automatically totals preceding entries. | Invoice templates that need a merged label cell and a dynamic total for line‑item amounts. | Automated monthly summaries that merge header cells and compute column totals without manual Excel editing.
// AI Prompts: Generate C# code that merges cells N10:O10, writes "Subtotal", and sets a SUM(N1:N9) formula in P10 with Aspose.Cells. | Explain step‑by‑step how to merge a range, add a label, apply a SUM formula, and save the workbook using Aspose.Cells for .NET. | Show an example of loading an Excel file, performing cell merging and formula insertion, then exporting the result with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates loading an Excel file with Aspose.Cells for .NET, merging cells N10:O10, inserting the text “Subtotal”, assigning a SUM(N1:N9) formula to P10, and saving the updated workbook.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells N10:O10 (zero‑based row 9, column 13, 1 row, 2 columns)
        cells.Merge(9, 13, 1, 2);

        // Put the subtotal label into the merged area
        cells["N10"].PutValue("Subtotal");

        // Calculate the sum of the adjacent column (column N, rows 1‑9) and place it in P10
        cells["P10"].Formula = "=SUM(N1:N9)";

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}
