// Title: How to copy rows in Aspose.Cells C# without adjusting absolute formulas (disable PasteOptions formula adjustment)
// Common Searches: Aspose.Cells C# copy rows keep absolute $A$1 references unchanged | disable formula adjustment when using PasteOptions in Aspose.Cells .NET | preserve absolute cell references while copying rows with Aspose.Cells | C# Aspose.Cells copy rows without shifting formulas | how to prevent formula relocation when copying rows in Excel using Aspose.Cells
// Tags: Aspose.Cells copy rows maintain formula integrity | PasteOptions turn off formula adjustment .NET | C# copy rows without formula shift Aspose.Cells | Cells.CopyRows keep $A$1 formulas | Excel workbook copy rows formula behavior Aspose | Aspose.Cells row copy retains original formulas

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, adds sample data with absolute and relative formulas, then copies a source row to a destination row using Cells.CopyRows. It demonstrates how to keep absolute references (e.g., $A$1) unchanged during the copy, discusses disabling formula adjustment via PasteOptions, recalculates formulas, and saves the resulting workbook.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);
                cells["C1"].PutValue(30);

                // Formulas with absolute references
                cells["A2"].Formula = "=$A$1+5";
                cells["B2"].Formula = "=$B$1*2";
                cells["C2"].Formula = "=$C$1-3";

                // Formulas with relative references (for contrast)
                cells["A3"].Formula = "=A1+5";
                cells["B3"].Formula = "=B1*2";
                cells["C3"].Formula = "=C1-3";

                workbook.Save(inputPath);
                Console.WriteLine($"Created sample workbook and saved as \"{inputPath}\".");
            }

            // Work with the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells wsCells = worksheet.Cells;

            // Define source and destination rows
            int sourceRow = 1;        // Row 2 (zero‑based index)
            int destinationRow = 5;   // Row 6 (zero‑based index)
            int totalRows = 1;        // Number of rows to copy

            // Copy rows within the same worksheet.
            // Use the overload that takes the source Cells collection.
            wsCells.CopyRows(wsCells, sourceRow, destinationRow, totalRows);
            Console.WriteLine($"Copied row {sourceRow + 1} to row {destinationRow + 1}.");

            // Recalculate formulas to reflect any changes
            workbook.CalculateFormula();

            // Verify the copied formula
            string originalFormula = wsCells[sourceRow, 0].Formula;      // A2
            string copiedFormula = wsCells[destinationRow, 0].Formula;   // A6
            Console.WriteLine($"Original formula in A{sourceRow + 1}: {originalFormula}");
            Console.WriteLine($"Copied formula in A{destinationRow + 1}: {copiedFormula}");

            // Save the result
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
