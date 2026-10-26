// Title: C# code to extract all formulas from an Excel workbook using Aspose.Cells and store them in a Dictionary keyed by SheetName!CellAddress
// Common Searches: how to get formula strings from each cell in an Excel file with Aspose.Cells C# | Aspose.Cells iterate worksheets and collect formulas into a dictionary | C# extract formulas from used range of workbook and map to sheet name and cell address | retrieve A1‑style formula references from Excel using Aspose.Cells API | list all cell formulas in a .xlsx file with Aspose.Cells for .NET
// Tags: Aspose.Cells extract workbook formulas C# | dictionary keyed by sheet name and cell address | enumerate used range cells Aspose.Cells | cell.IsFormula usage Aspose.Cells | load Excel file and read formula strings .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet's used range, checks each cell for a formula via the IsFormula property, and stores the formula text in a Dictionary where the key combines the worksheet name and the cell's A1 address.
public class FormulaExtractor
{
    // Extracts all formulas from every worksheet in the workbook.
    // Returns a dictionary where the key is "SheetName!CellAddress" (e.g., "Sheet1!B2")
    // and the value is the formula string (e.g., "=SUM(A1:A10)").
    public static Dictionary<string, string> ExtractFormulas(string filePath)
    {
        var formulas = new Dictionary<string, string>();

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Scan each cell within the used range.
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell actually contains a formula.
                        if (cell != null && cell.IsFormula)
                        {
                            // Build a unique key using sheet name and cell address (A1 style).
                            string key = $"{sheet.Name}!{cell.Name}";

                            // Store the formula text.
                            formulas[key] = cell.Formula;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting formulas: {ex.Message}");
        }

        return formulas;
    }
}

public class Program
{
    // Entry point for the console application.
    public static void Main(string[] args)
    {
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        // Ensure the input file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Extract formulas from the workbook.
            Dictionary<string, string> formulas = FormulaExtractor.ExtractFormulas(filePath);

            // Output the extracted formulas.
            foreach (KeyValuePair<string, string> kvp in formulas)
            {
                Console.WriteLine($"{kvp.Key} = {kvp.Value}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
