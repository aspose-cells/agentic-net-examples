// Title: Generate a console report of zero‑based row and column indices for cells with Style.QuotePrefix = true using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates a worksheet and prints the zero‑based row and column numbers of every cell whose Style.QuotePrefix property is true. | Modify the example to output the Excel address (e.g., B10) instead of numeric indices for cells that have QuotePrefix enabled. | Extend the program to save the list of QuotePrefix cells to a CSV file, including row index, column index, and cell address.
// Common Searches: aspnet cells enumerate cells with quote prefix property true | c# aspose.cells find cells that have leading apostrophe flag | how to list rows and columns of quoted cells in Excel using Aspose.Cells | retrieve zero based indices of cells with Style.QuotePrefix in .NET | generate report of cells with QuotePrefix using Aspose.Cells library
// Tags: Aspose.Cells enumerate QuotePrefix cells | C# generate console report of quoted Excel cells | export quoted cell indices to CSV using Aspose.Cells | zero‑based row and column extraction with Style.QuotePrefix | list cells with leading apostrophe in .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace QuotePrefixReport
{
    // Creates a workbook, sets QuotePrefix on sample cells, scans all used cells, collects zero‑based row and column indices of cells where Style.QuotePrefix is true, prints the list to the console, and saves the workbook as QuotePrefixReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: set some cells with QuotePrefix = true
            Cell cellB10 = cells["B10"];
            cellB10.PutValue("'12345");               // value with leading quote
            Style styleB10 = cellB10.GetStyle();
            styleB10.QuotePrefix = true;              // enable QuotePrefix
            cellB10.SetStyle(styleB10);

            Cell cellC5 = cells["C5"];
            cellC5.PutValue("'Hello");
            Style styleC5 = cellC5.GetStyle();
            styleC5.QuotePrefix = true;
            cellC5.SetStyle(styleC5);

            // List to hold report lines
            List<string> reportLines = new List<string>();
            reportLines.Add("Cells with QuotePrefix = true:");
            reportLines.Add("Row\tColumn");

            // Iterate through all used cells in the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell curCell = cells[row, col];
                    // Skip empty cells
                    if (curCell == null || curCell.Type == CellValueType.IsNull) continue;

                    // Check the QuotePrefix property
                    if (curCell.GetStyle().QuotePrefix)
                    {
                        // Add row and column indices (zero‑based) to the report
                        reportLines.Add($"{row}\t{col}");
                    }
                }
            }

            // Output the report to the console
            foreach (string line in reportLines)
            {
                Console.WriteLine(line);
            }

            // Optionally, save the workbook (lifecycle rule)
            workbook.Save("QuotePrefixReport.xlsx");
        }
    }
}
