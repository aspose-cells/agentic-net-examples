// Title: Change a cell reference in an Excel formula and recalculate the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Load input.xlsx using Aspose.Cells, replace the reference A1 with A3 in the formula of cell B2, invoke CalculateFormula, and save the result as output.xlsx. | Open a workbook, edit the Formula property of a specific cell, trigger a full recalculation, then read the updated Value property in C# with Aspose.Cells.
// Common Searches: Aspose.Cells C# edit formula to point to another cell and recalculate workbook | replace cell reference in Excel formula programmatically using Aspose.Cells .NET | how to update an Excel formula and get the new value with Aspose.Cells in C# | calculate updated formula result after changing cell reference in Aspose.Cells
// Tags: modify cell formula Aspose.Cells C# | calculate workbook formulas Aspose.Cells | change cell reference in Excel formula .NET | update formula and recalculate Aspose.Cells | load edit save workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example ensures an input.xlsx exists (creating a simple workbook if needed), loads it with Aspose.Cells, changes the formula in cell B2 from referencing A1 to A3, recalculates all formulas, prints the updated value, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a simple workbook for demonstration.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                // Populate cells A1 and A3 with sample numbers.
                tempSheet.Cells["A1"].PutValue(10);
                tempSheet.Cells["A3"].PutValue(30);
                // Place a formula in B2 that references A1.
                tempSheet.Cells["B2"].Formula = "=A1*2";
                tempWb.Save(inputPath);
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Identify the cell that contains the formula (example: B2).
            Cell formulaCell = sheet.Cells["B2"];

            // Verify that the cell actually contains a formula.
            if (string.IsNullOrEmpty(formulaCell.Formula))
            {
                Console.WriteLine($"Cell {formulaCell.Name} does not contain a formula.");
                return;
            }

            // Store the original formula.
            string originalFormula = formulaCell.Formula;

            // Modify the formula to reference a different cell.
            // Example: change reference from A1 to A3 while keeping the rest of the formula unchanged.
            string modifiedFormula = originalFormula.Replace("A1", "A3");
            formulaCell.Formula = modifiedFormula;

            // Recalculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Retrieve and display the updated value of the modified cell.
            object updatedValue = formulaCell.Value;
            Console.WriteLine($"Updated value in {formulaCell.Name}: {updatedValue}");

            // Save the workbook with the changes.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
