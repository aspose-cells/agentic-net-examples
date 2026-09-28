// Title: How to enumerate all cells in an Aspose.Cells worksheet and count empty, string, numeric, date, boolean, and formula cells in C#
// AI Prompts: Generate C# code using Aspose.Cells that loops through every instantiated cell in a worksheet and returns a dictionary with counts for empty, string, numeric, date, boolean, and formula cells. | Extend the cell‑type counting loop to calculate the percentage of each type relative to the total cell count and display the results.
// Common Searches: Aspose.Cells C# count empty cells and different data types in a worksheet | C# enumerate worksheet cells with Aspose.Cells and get statistics per cell type | How to get number of formula cells using Aspose.Cells .NET | Aspose.Cells get count of string numeric date boolean cells in Excel file | C# Aspose.Cells iterate over cells and differentiate null and empty string
// Tags: enumerate worksheet cells Aspose.Cells C# | count cell types Aspose.Cells | empty cell detection Aspose.Cells | formula cell counting Aspose.Cells | cell type statistics Aspose.Cells XLSX

using System;
using System.Collections;
using Aspose.Cells;

// The example creates a workbook, adds sample data of various types (string, numeric, date, boolean, null, empty string, formula), iterates over all instantiated cells with Aspose.Cells, counts total, empty, string, numeric, date, boolean, and formula cells, outputs the counts, and saves the file as CellSummary.xlsx.
class EmptyFilledCellSummary
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data with different types and some empty cells
            cells["A1"].PutValue("Hello");               // string
            cells["B1"].PutValue(123);                   // numeric
            cells["C1"].PutValue(DateTime.Now);          // date
            cells["D1"].PutValue(true);                  // boolean
            cells["E1"].PutValue(null);                  // empty (null)
            cells["F1"].PutValue(string.Empty);          // empty (empty string)
            cells["G1"].Formula = "=SUM(B1)";            // formula

            // Counters for summary
            long totalCells = 0;
            long emptyCells = 0;
            long stringCells = 0;
            long numericCells = 0;
            long dateCells = 0;
            long booleanCells = 0;
            long formulaCells = 0;

            // Enumerate all instantiated cells in the worksheet
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                totalCells++;

                // Determine if the cell is empty (null or CellValueType.IsNull)
                if (cell.Type == CellValueType.IsNull ||
                    cell.Value == null ||
                    (cell.Value is string s && string.IsNullOrEmpty(s)))
                {
                    emptyCells++;
                    continue;
                }

                // Count based on cell content type
                if (cell.IsFormula)
                {
                    formulaCells++;
                }
                else if (cell.Value is string)
                {
                    stringCells++;
                }
                else if (cell.Value is double ||
                         cell.Value is int ||
                         cell.Value is decimal ||
                         cell.Value is float)
                {
                    numericCells++;
                }
                else if (cell.Value is DateTime)
                {
                    dateCells++;
                }
                else if (cell.Value is bool)
                {
                    booleanCells++;
                }
            }

            // Output the summary
            Console.WriteLine($"Total instantiated cells: {totalCells}");
            Console.WriteLine($"Empty cells: {emptyCells}");
            Console.WriteLine($"String cells: {stringCells}");
            Console.WriteLine($"Numeric cells: {numericCells}");
            Console.WriteLine($"Date cells: {dateCells}");
            Console.WriteLine($"Boolean cells: {booleanCells}");
            Console.WriteLine($"Formula cells: {formulaCells}");

            // Save the workbook
            workbook.Save("CellSummary.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
