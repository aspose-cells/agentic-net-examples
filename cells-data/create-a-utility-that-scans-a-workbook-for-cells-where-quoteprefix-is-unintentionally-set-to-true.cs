// Title: C# utility to locate Excel cells where the QuotePrefix style flag is true using Aspose.Cells
// Common Searches: how to programmatically detect cells with QuotePrefix set in an Excel file using Aspose.Cells C# | C# code to scan a workbook and list addresses of cells that have QuotePrefix enabled | Aspose.Cells find cells with unintended QuotePrefix property in .NET | retrieve cell references with QuotePrefix true from an Excel workbook in C#
// Tags: Aspose.Cells detect QuotePrefix flag | C# scan Excel for QuotePrefix cells | list cell addresses with QuotePrefix true | Excel style QuotePrefix detection utility | identify unintended QuotePrefix in worksheets

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The QuotePrefixScanner class loads an Excel workbook, iterates through each worksheet's used range, checks each cell's style for the QuotePrefix flag, and collects the sheet name and cell address of every cell where QuotePrefix is true, returning the list for further processing or reporting.
class QuotePrefixScanner
{
    // Scans the workbook at the given path and returns a list of cell addresses
    // where the QuotePrefix property is set to true.
    public static List<string> FindQuotePrefixCells(string filePath)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(filePath);
        List<string> cellsWithQuotePrefix = new List<string>();

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range of rows and columns.
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Loop through all cells within the used range.
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell == null) continue;

                    // Access the cell's style and check the QuotePrefix flag.
                    if (cell.GetStyle().QuotePrefix)
                    {
                        // Record the cell address with its sheet name.
                        cellsWithQuotePrefix.Add($"{sheet.Name}!{cell.Name}");
                    }
                }
            }
        }

        return cellsWithQuotePrefix;
    }

    // Example entry point: accepts a file path argument and prints the results.
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: QuotePrefixScanner <excel-file-path>");
            return;
        }

        string filePath = args[0];
        List<string> quotePrefixCells = FindQuotePrefixCells(filePath);

        Console.WriteLine("Cells with QuotePrefix set to true:");
        foreach (string address in quotePrefixCells)
        {
            Console.WriteLine(address);
        }
    }
}
