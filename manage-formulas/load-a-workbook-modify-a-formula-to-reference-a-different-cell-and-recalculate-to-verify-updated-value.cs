// Title: Change a cell formula reference and recalculate workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to modify a formula in an existing Excel file with Aspose.Cells C# | replace A1 with A2 in a worksheet formula using Aspose.Cells .NET | recalculate workbook after updating a cell formula programmatically Aspose.Cells | load workbook, edit cell B2 formula, and save changes with Aspose.Cells C# example | Aspose.Cells calculate all formulas after changing a formula reference
// Tags: modify cell formula Aspose.Cells | recalculate all formulas Aspose.Cells | replace Excel cell reference C# | load and save workbook Aspose.Cells .NET | programmatic formula update Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing 'input.xlsx' workbook (or creates one), accesses cell B2, replaces the formula reference from A1 to A2, triggers a full workbook recalculation, outputs the new cell value, and saves the updated workbook as 'output.xlsx', handling missing cells and errors.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one with sample data.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Cells["A1"].PutValue(10);
                    ws.Cells["B2"].Formula = "A1";
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Access the target cell (B2) that contains the formula.
                Cell formulaCell = worksheet.Cells["B2"];
                if (formulaCell == null)
                {
                    Console.WriteLine("Cell B2 not found.");
                    return;
                }

                // Retrieve the original formula; handle empty or null formulas safely.
                string originalFormula = formulaCell.Formula ?? string.Empty;
                if (string.IsNullOrEmpty(originalFormula))
                {
                    Console.WriteLine("Cell B2 does not contain a formula to modify.");
                    return;
                }

                // Replace reference "A1" with "A2" (adjust as needed for more complex scenarios).
                string updatedFormula = originalFormula.Replace("A1", "A2");
                formulaCell.Formula = updatedFormula;

                // Recalculate all formulas to reflect the change.
                workbook.CalculateFormula();

                // Output the new value of the modified cell.
                object newValue = formulaCell.Value;
                Console.WriteLine($"Cell {formulaCell.Name} updated value: {newValue}");

                // Save the workbook with the modifications.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
