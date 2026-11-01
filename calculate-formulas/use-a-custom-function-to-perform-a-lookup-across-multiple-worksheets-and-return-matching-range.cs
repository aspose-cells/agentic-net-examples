// Title: C# Aspose.Cells example: perform a multi‑sheet lookup and get the cell address with sheet name
// Common Searches: Aspose.Cells C# find first occurrence of a value across all worksheets | How to return sheet name and cell reference when searching an Excel workbook with Aspose.Cells | C# code to search every worksheet for a specific value and output the address | Multi‑sheet lookup using Aspose.Cells API in .NET | Retrieve matching cell address from multiple sheets in an Excel file with Aspose.Cells
// Tags: multi-sheet lookup Aspose.Cells C# | return cell address with sheet name Aspose.Cells | search all worksheets for value .NET | iterate workbook cells Aspose.Cells | calculate formulas after lookup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample defines a MultiSheetLookup method that iterates through each worksheet, row, and column of an Aspose.Cells Workbook, compares cell values to a given lookup string, and returns the first matching cell address prefixed with its sheet name (e.g., 'Sheet1'!B2) or "#N/A" if not found. The Main routine loads or creates a workbook, inserts a test value, calls the lookup, writes the result to cell A1 of the first sheet, recalculates formulas, and saves the file.
public class Program
{
    // Performs a lookup across all worksheets and returns the first matching cell address.
    private static string MultiSheetLookup(Workbook workbook, string lookupValue)
    {
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Cells cells = ws.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int r = 0; r <= maxRow; r++)
            {
                for (int c = 0; c <= maxCol; c++)
                {
                    Cell cell = cells[r, c];
                    if (cell.Value != null && cell.Value.ToString() == lookupValue)
                    {
                        // Return address with sheet name, e.g., 'Sheet1'!B2
                        return $"'{ws.Name}'!{cell.Name}";
                    }
                }
            }
        }

        return "#N/A";
    }

    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string searchValue = "SearchValue";

        Workbook workbook = null;

        try
        {
            // Ensure the input file exists; if not, create a simple workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["B2"].PutValue(searchValue); // Sample data for lookup
            }

            // Perform the multi‑sheet lookup.
            string result = MultiSheetLookup(workbook, searchValue);

            // Place the result in A1 of the first worksheet.
            Worksheet firstSheet = workbook.Worksheets[0];
            Cell targetCell = firstSheet.Cells["A1"];
            targetCell.PutValue(result);

            // Calculate formulas (if any) and save the workbook.
            workbook.CalculateFormula();
            workbook.Save(outputPath);
            Console.WriteLine("Lookup result: " + result);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
