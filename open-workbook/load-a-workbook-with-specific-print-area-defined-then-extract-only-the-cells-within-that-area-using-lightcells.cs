// Title: Extract cells within a worksheet's defined print area using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read only the print area of an Excel worksheet | how to get cells from a print area range with Aspose.Cells .NET | extract values from a defined print area in an .xlsx file using Aspose.Cells | parse PageSetup.PrintArea string to CellArea in C# Aspose.Cells | iterate over cells in a print area without loading entire sheet Aspose.Cells
// Tags: Aspose.Cells extract print area cells | C# parse print area string CellArea | Aspose.Cells PageSetup print area handling | Excel workbook load specific range Aspose.Cells | iterate worksheet cells within defined print area

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, retrieves the worksheet's print area from PageSetup, converts the range string into a CellArea object, and iterates over each cell in that area, outputting addresses and values while handling missing files, undefined print areas, and parsing errors.
class ExtractPrintArea
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (the file must have a print area defined)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the print area string (e.g., "A1:C10")
            string printAreaRef = sheet.PageSetup.PrintArea;

            // If no print area is defined, exit
            if (string.IsNullOrEmpty(printAreaRef))
            {
                Console.WriteLine("No print area defined in the worksheet.");
                return;
            }

            // Parse the print area reference into a CellArea structure
            CellArea printArea;
            try
            {
                // Print area can be a single cell or a range like "A1:C10"
                string[] parts = printAreaRef.Split(':');
                if (parts.Length == 2)
                {
                    CellsHelper.CellNameToIndex(parts[0], out int startRow, out int startColumn);
                    CellsHelper.CellNameToIndex(parts[1], out int endRow, out int endColumn);
                    printArea = new CellArea
                    {
                        StartRow = startRow,
                        StartColumn = startColumn,
                        EndRow = endRow,
                        EndColumn = endColumn
                    };
                }
                else
                {
                    // Single cell reference
                    CellsHelper.CellNameToIndex(printAreaRef, out int row, out int column);
                    printArea = new CellArea
                    {
                        StartRow = row,
                        StartColumn = column,
                        EndRow = row,
                        EndColumn = column
                    };
                }
            }
            catch (Exception exParse)
            {
                Console.WriteLine($"Failed to parse the print area reference: {exParse.Message}");
                return;
            }

            // Iterate through the cells within the print area
            for (int row = printArea.StartRow; row <= printArea.EndRow; row++)
            {
                for (int col = printArea.StartColumn; col <= printArea.EndColumn; col++)
                {
                    try
                    {
                        Cell cell = sheet.Cells[row, col];
                        string address = CellsHelper.CellIndexToName(row, col);
                        Console.WriteLine($"{address}: {cell.Value}");
                    }
                    catch (Exception exCell)
                    {
                        Console.WriteLine($"Error reading cell at row {row}, column {col}: {exCell.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
