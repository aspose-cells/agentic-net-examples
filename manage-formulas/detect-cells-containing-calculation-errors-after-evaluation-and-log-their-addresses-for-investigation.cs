// Title: Identify and log Excel cells that contain calculation errors after evaluating formulas with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, calls Workbook.CalculateFormula, scans all worksheets, and returns a list of A1‑style addresses for cells where Cell.Type equals CellValueType.IsError. | Modify the example to export the collected error cell addresses to a CSV file, including the worksheet name as a prefix for each entry.
// Common Searches: Aspose.Cells .NET find cells that return #VALUE! after CalculateFormula | C# list all error cells in an Excel workbook using Aspose.Cells | how to extract addresses of formula errors from a spreadsheet with Aspose.Cells | detect #DIV/0 and other calculation errors in Excel using Aspose.Cells C# | log cells with evaluation errors after running Workbook.CalculateFormula in .NET
// Tags: detect calculation errors Aspose.Cells | collect error cell addresses C# | scan used range for IsError Aspose.Cells | export error cell list to CSV .NET | Workbook.CalculateFormula error handling

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The sample loads an Excel workbook, evaluates all formulas, iterates each worksheet's used range, records the A1‑style addresses of cells whose value type is IsError, and outputs the list of error cell addresses for further investigation.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Evaluate all formulas in the workbook
        workbook.CalculateFormula();

        // List to hold addresses of cells with errors
        List<string> errorCellAddresses = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a calculation error
                    if (cell.Type == CellValueType.IsError)
                    {
                        // Record the address in A1 style, prefixed with the sheet name
                        string address = $"{sheet.Name}!{cell.Name}";
                        errorCellAddresses.Add(address);
                    }
                }
            }
        }

        // Output the addresses of all error cells
        Console.WriteLine("Cells containing calculation errors:");
        foreach (string addr in errorCellAddresses)
        {
            Console.WriteLine(addr);
        }

        // Optional: save the workbook if modifications were made
        // workbook.Save("output.xlsx");
    }
}
