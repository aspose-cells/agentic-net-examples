// Title: How to merge cells N10:O10, add a "Subtotal" label, and insert a SUM formula with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge N10 O10 and set subtotal label | C# code to add SUM(N1:N9) formula after merging cells with Aspose.Cells | Create or load an Excel workbook and merge cells using Aspose.Cells .NET | Insert a subtotal row in Excel programmatically with Aspose.Cells C# | How to programmatically merge cells and apply a formula in Aspose.Cells for .NET
// Tags: merge cells N10 O10 Aspose.Cells | add subtotal label merged cells C# | insert SUM(N1:N9) formula Aspose.Cells | conditional workbook load or create Aspose.Cells | save Excel file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, merges cells N10 and O10 on the first worksheet, writes the text "Subtotal" into the merged range, places a SUM formula in N11 that totals N1 through N9, and saves the workbook to the specified output file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Merge cells N10:O10 (zero‑based indices: row 9, column 13)
                sheet.Cells.Merge(9, 13, 1, 2);

                // Write the subtotal label into the merged range
                sheet.Cells["N10"].PutValue("Subtotal");

                // Place a SUM formula in the cell below (N11) that adds the values of column N above the label
                sheet.Cells["N11"].Formula = "=SUM(N1:N9)";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
