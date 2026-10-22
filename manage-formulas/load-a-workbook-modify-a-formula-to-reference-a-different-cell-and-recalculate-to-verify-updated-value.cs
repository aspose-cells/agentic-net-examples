// Title: Change a cell reference in an Excel formula and recalculate the workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace cell reference in existing formula | How to update an Excel formula to use a different cell and recalculate with Aspose.Cells | Recalculate workbook after modifying a formula programmatically in .NET | Load workbook and edit B2 formula from A1 to A2 using Aspose.Cells | Get updated formula result after changing cell reference in Aspose.Cells
// Tags: Aspose.Cells modify formula reference | recalculate workbook formulas .NET | programmatic Excel formula update C# | change cell address in Aspose.Cells formula | get updated formula value Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook (or creates one), changes the formula in cell B2 from referencing A1 to referencing A2, triggers a full recalculation of all formulas, prints the new value of the modified cell, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a sample workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue(10);
                ws.Cells["B2"].Formula = "=A1*2";
            }

            Worksheet worksheet = workbook.Worksheets[0];
            Cell formulaCell = worksheet.Cells["B2"];

            // Modify the formula only if it exists
            if (!string.IsNullOrEmpty(formulaCell.Formula))
            {
                string originalFormula = formulaCell.Formula;
                string updatedFormula = originalFormula.Replace("A1", "A2");
                formulaCell.Formula = updatedFormula;
            }

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Display the updated value
            Console.WriteLine($"Updated value in {formulaCell.Name}: {formulaCell.Value}");

            // Save the workbook with changes
            workbook.Save(outputPath);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
