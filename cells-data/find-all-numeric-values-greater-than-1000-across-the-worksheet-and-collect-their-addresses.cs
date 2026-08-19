// Title: Aspose.Cells C# example: List addresses of cells with values > 1000 and export to a new worksheet
// Description: Loads an Excel file, scans the used range of the first worksheet, captures the addresses of cells whose numeric value (int, double, or parsable string) exceeds 1,000, writes those addresses to a newly created sheet named "Values>1000", and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | Excel | find cells greater than 1000 | cell address list | used range iteration | export to new worksheet | threshold filter | sample code
// Common Searches: Aspose.Cells find cells with value over 1000 | C# list cell addresses exceeding a threshold | How to export high‑value cells to another sheet using Aspose.Cells | Iterate used range in Excel with Aspose.Cells .NET
// Developer Intent: Extract every cell address whose numeric content is larger than a specified limit and store the results in a separate worksheet.
// Use Cases: Create an audit sheet of outlier amounts in financial reports | Pre‑process data by flagging entries that surpass a business rule | Generate a quick summary of high‑value items for dashboards
// AI Prompts: Write a reusable method that returns cell addresses where the value exceeds a given threshold using Aspose.Cells. | Refactor the nested loops into a LINQ query or use Aspose.Cells' FindAll to locate cells > 1000. | Add error handling that logs non‑numeric cells and continues processing without interruption.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel file, scans the used range of the first worksheet, captures the addresses of cells whose numeric value (int, double, or parsable string) exceeds 1,000, writes those addresses to a newly created sheet named "Values>1000", and saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // List to store addresses of cells whose numeric value is greater than 1000
        List<string> addresses = new List<string>();

        // Determine the used range of the worksheet
        int maxRow = worksheet.Cells.MaxDataRow;
        int maxCol = worksheet.Cells.MaxDataColumn;

        // Scan each cell within the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = worksheet.Cells[row, col];
                if (cell == null || cell.Value == null)
                    continue;

                // Direct numeric types
                if (cell.Value is double d && d > 1000)
                {
                    addresses.Add(cell.Name);
                }
                else if (cell.Value is int i && i > 1000)
                {
                    addresses.Add(cell.Name);
                }
                else
                {
                    // Attempt to parse string representations of numbers
                    if (double.TryParse(cell.StringValue, out double parsed) && parsed > 1000)
                    {
                        addresses.Add(cell.Name);
                    }
                }
            }
        }

        // Write the collected addresses to a new worksheet
        int resultIndex = workbook.Worksheets.Add();
        Worksheet resultSheet = workbook.Worksheets[resultIndex];
        resultSheet.Name = "Values>1000";

        resultSheet.Cells[0, 0].PutValue("CellAddress");
        for (int i = 0; i < addresses.Count; i++)
        {
            resultSheet.Cells[i + 1, 0].PutValue(addresses[i]);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
