// Title: Locate cells with Excel error values (e.g., #DIV/0!) inside a named range using Aspose.Cells for .NET
// Common Searches: how to get addresses of cells that contain #DIV/0! in a named range with Aspose.Cells C# | Aspose.Cells C# iterate over a named range and identify cells with error values | retrieve error cells from a specific named range using Aspose.Cells for .NET | C# find cells where CellValueType.IsError is true inside a workbook named range | search for Excel error values in a defined range with Aspose.Cells API
// Tags: Aspose.Cells locate error cells in named range | C# iterate named range error values Aspose.Cells | Excel error value detection Aspose.Cells .NET | retrieve cell addresses with IsError Aspose.Cells | named range cell error extraction C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, obtains a named range, loops through each cell in that range, checks for error values (such as #DIV/0!) using the CellValueType.IsError flag, collects the cell addresses, and prints the list of error cells.
class FindErrorCellsInNamedRange
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string rangeName = "MyRange";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            AsposeRange range = workbook.Worksheets.GetRangeByName(rangeName);
            if (range == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // List to hold addresses of cells containing error values
            List<string> errorCells = new List<string>();

            // Iterate through all cells in the range
            int startRow = range.FirstRow;
            int startColumn = range.FirstColumn;
            int rowCount = range.RowCount;
            int columnCount = range.ColumnCount;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = range.Worksheet.Cells[startRow + i, startColumn + j];

                    // Check if the cell contains an error value
                    if (cell.Type == CellValueType.IsError)
                    {
                        // Collect the cell address (e.g., "A1")
                        errorCells.Add(cell.Name);
                    }
                }
            }

            // Output the results
            Console.WriteLine($"Cells with error values in named range \"{rangeName}\":");
            foreach (string address in errorCells)
            {
                Console.WriteLine(address);
            }

            // (Optional) Save the workbook if modifications are made
            // workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
