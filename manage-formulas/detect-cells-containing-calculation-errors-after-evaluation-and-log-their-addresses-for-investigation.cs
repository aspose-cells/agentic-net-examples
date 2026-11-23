// Title: Identify and log Excel cells that return calculation errors after formula recalculation using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells find cells that contain #DIV/0! after CalculateFormula | how to log addresses of error cells in an Excel workbook using Aspose.Cells .NET | detect formula evaluation errors in each worksheet with Aspose.Cells API | iterate used range and check CellValueType.IsError in Aspose.Cells C# | Aspose.Cells .NET retrieve cells with error values after recalculating formulas
// Tags: Aspose.Cells error cell detection | C# used range iteration Aspose.Cells | CellValueType.IsError usage | Excel formula recalculation error capture | log error cell addresses Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsErrorDetection
{
    // The example loads an Excel workbook, recalculates all formulas, then scans each worksheet's used range. It checks each cell's type for errors (CellValueType.IsError) and writes the sheet name, cell address, and error message to the console for further investigation.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell contains a calculation error
                        if (cell.Type == CellValueType.IsError)
                        {
                            // Log the address and error value
                            Console.WriteLine(
                                $"Error detected in sheet '{sheet.Name}' at {cell.Name}: {cell.StringValue}");
                        }
                    }
                }
            }

            // Optional: Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}
