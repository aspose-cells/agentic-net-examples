// Title: Force worksheet formula recalculation after deleting rows or columns with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# recalculate formulas after deleting a row | how to trigger worksheet calculation after DeleteColumn in Aspose.Cells | C# delete multiple rows and update formulas with Aspose.Cells | force workbook formula evaluation after removing columns using Aspose.Cells .NET
// Tags: delete row worksheet.CalculateFormula Aspose.Cells | column deletion formula refresh C# Aspose.Cells | batch row removal recalculate workbook Aspose.Cells | force formula evaluation after DeleteRows Aspose.Cells | Excel file manipulation recalculate formulas Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, deletes a specific row, a column, and a range of rows, and calls CalculateFormula after each deletion to ensure all formulas are recomputed before saving the updated file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFile);
            Worksheet ws = workbook.Worksheets[0];

            // Delete row 2 (zero‑based index 1) and recalculate
            ws.Cells.DeleteRow(1);
            workbook.CalculateFormula();

            // Delete column 3 (zero‑based index 2) and recalculate
            ws.Cells.DeleteColumn(2);
            workbook.CalculateFormula();

            // Delete rows 2 to 4 (start index 1, total 3 rows) and recalculate
            ws.Cells.DeleteRows(1, 3);
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
