// Title: C# loop through only populated cells in an Aspose.Cells worksheet using MaxDataRow and MaxDataColumn
// AI Prompts: Write a C# routine that uses Workbook.Worksheets[0].Cells.MaxDataRow and MaxDataColumn to traverse only cells that contain data and print each cell's address and value. | Show how to guard against an empty worksheet and skip null cells while iterating a bounded range in Aspose.Cells with C#. | Create a C# example that saves the workbook after enumerating non‑empty cells within the detected data rectangle.
// Common Searches: aspnet c# iterate over non empty cells Aspose.Cells MaxDataRow | how to get last used row and column in Aspose.Cells C# | skip blank rows when reading Excel with Aspose.Cells C# | process only populated range in Aspose.Cells workbook using MaxDataColumn | C# Aspose.Cells loop through cells without checking every row
// Tags: Aspose.Cells iterate populated range C# | MaxDataRow MaxDataColumn usage Aspose.Cells | skip empty rows Aspose.Cells C# | process non‑empty cells Aspose.Cells workbook | retrieve last data cell Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsMaxDataIteration
{
    // The sample creates a workbook, adds data with intentional empty rows and columns, obtains the last populated row and column via MaxDataRow and MaxDataColumn, and then iterates over that bounded range, processing only instantiated cells that hold values before saving the result to an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with some empty rows/columns
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["A2"].PutValue("Item1");
            cells["B2"].PutValue(10);
            // Row 3 is intentionally left empty
            cells["A4"].PutValue("Item2");
            cells["B4"].PutValue(20);
            // Column C is empty

            // Retrieve the limits of populated area (property rules)
            int maxRow = cells.MaxDataRow;       // zero‑based index of last row containing data
            int maxCol = cells.MaxDataColumn;    // zero‑based index of last column containing data

            // Guard against empty worksheet
            if (maxRow >= 0 && maxCol >= 0)
            {
                Console.WriteLine($"Iterating over populated range: Rows 0‑{maxRow}, Columns 0‑{maxCol}");
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Get the cell; may be null if never instantiated
                        Cell cell = cells[row, col];
                        if (cell != null && cell.Value != null)
                        {
                            // Process the cell (example: output its address and value)
                            Console.WriteLine($"{cell.Name}: {cell.Value}");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Worksheet contains no data.");
            }

            // Save the workbook (save rule)
            workbook.Save("MaxDataIterationResult.xlsx");
        }
    }
}
