// Title: Find and report Excel formulas that return #REF! after rows are deleted using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# detect #REF! errors caused by deleted rows | list cells with broken references after removing rows using Aspose.Cells | how to programmatically identify formulas that reference missing rows in .NET | C# example to scan workbook for #REF! formula errors with Aspose.Cells | detect invalid cell references after row deletion in Excel via Aspose.Cells API
// Tags: Aspose.Cells detect #REF! formulas | C# scan Excel workbook for broken references | identify deleted row references in Excel formulas | Aspose.Cells iterate worksheets cells errors | handle #REF! errors with Aspose.Cells .NET | Excel formula validation using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file with Aspose.Cells, iterates through all worksheets and cells, checks each formula for a #REF! result caused by deleted rows or columns, and prints the sheet name, cell address, offending formula, and a suggestion to replace the invalid reference.
class FormulaReferenceChecker
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Determine if the cell contains a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Check if the formula result is a #REF! error
                        // When a formula evaluates to an error, the cell's value is a string like "#REF!"
                        if (cell.Value != null && cell.Value.ToString().Equals("#REF!", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"Sheet: {sheet.Name}, Cell: {cell.Name}");
                            Console.WriteLine($"  Formula: {cell.Formula}");
                            Console.WriteLine("  Issue: Formula references a deleted row/column (produces #REF! error).");
                            Console.WriteLine("  Suggestion: Review the formula and replace the #REF! reference with a valid cell/range.");
                            Console.WriteLine();
                        }
                    }
                }
            }

            // Optionally, save the workbook after manual corrections (uncomment if needed)
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
