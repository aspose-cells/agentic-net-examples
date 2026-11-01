// Title: Log addresses of empty cells while enumerating all cells in a worksheet using Aspose.Cells for .NET
// AI Prompts: Write a C# program that uses Aspose.Cells to iterate every cell in the first worksheet, checks for null values, and writes each empty cell's address to a log file. | Create a .NET snippet that scans an Excel sheet with Aspose.Cells, captures the addresses of cells where Value is null, and saves the list to a text file without modifying the workbook.
// Common Searches: asp.net find empty cells in Excel workbook using Aspose.Cells | how to export list of blank cell addresses from Excel with C# | Aspose.Cells enumerate cells and log null values to file | detect and record empty cells for data quality audit in .NET | save workbook after scanning for null cells Aspose.Cells
// Tags: enumerate worksheet cells Aspose.Cells | detect null cell values C# | export empty cell addresses to text file | Aspose.Cells data completeness validation | log blank cells in Excel .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates all cells in the first worksheet, identifies cells whose Value is null, writes their addresses to a text log, and saves the workbook unchanged.
class NullCellLogger
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Store addresses of cells that contain null values
        List<string> nullCellAddresses = new List<string>();

        // Enumerate all cells in the worksheet
        foreach (Cell cell in cells)
        {
            // Check if the cell's value is null (empty)
            if (cell.Value == null)
            {
                // Log the cell address (e.g., "A1")
                string address = cell.Name;
                Console.WriteLine($"Null value found at {address}");
                nullCellAddresses.Add(address);
            }
        }

        // Write the list of null cell addresses to a log file
        System.IO.File.WriteAllLines("NullCellsLog.txt", nullCellAddresses);

        // Save the workbook (no modifications made, but lifecycle rule requires a save)
        workbook.Save("output.xlsx");
    }
}
