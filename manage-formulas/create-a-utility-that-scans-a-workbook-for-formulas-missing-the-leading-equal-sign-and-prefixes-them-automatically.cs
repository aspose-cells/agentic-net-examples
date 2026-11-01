// Title: C# utility to prepend missing leading '=' to Excel formulas using Aspose.Cells
// Common Searches: C# Aspose.Cells add missing equal sign to cell formulas in an existing workbook | how to fix formulas without leading '=' in Excel files with Aspose.Cells .NET | scan all worksheets and prepend '=' to formulas that lack it using Aspose.Cells | bulk correct Excel formulas missing '=' with Aspose.Cells in a console application | automate formula normalization for Excel workbooks using Aspose.Cells library
// Tags: prepend missing equal sign to Excel formulas Aspose.Cells | detect formulas lacking leading '=' .NET | bulk update cell formulas Aspose.Cells C# | iterate worksheets and cells Aspose.Cells | normalize Excel formula syntax programmatically

using Aspose.Cells;
using System;

// The example loads a workbook, iterates through each worksheet's used range, checks each cell's Formula property, and if the formula does not start with '=', it prefixes the missing equal sign before saving the corrected workbook.
class FormulaFixer
{
    static void Main(string[] args)
    {
        // Expect input and output file paths as arguments
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: FormulaFixer <input.xlsx> <output.xlsx>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Process each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to avoid scanning empty rows/columns
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Skip cells without a formula
                    if (cell == null || string.IsNullOrEmpty(cell.Formula))
                        continue;

                    string formula = cell.Formula;

                    // If the formula does not start with '=', it is missing the leading equal sign
                    if (!formula.StartsWith("="))
                    {
                        // Prefix the missing '=' and assign back as a proper formula
                        cell.Formula = "=" + formula;
                    }
                }
            }
        }

        // Save the corrected workbook
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
