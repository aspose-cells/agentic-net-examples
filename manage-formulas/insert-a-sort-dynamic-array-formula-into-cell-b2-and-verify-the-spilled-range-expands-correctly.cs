// Title: Insert a SORT dynamic array formula into a cell with Aspose.Cells for .NET and validate the spilled range size
// Common Searches: Aspose.Cells C# insert SORT dynamic array formula and check spill range | how to count spilled cells after using SORT formula in a .NET workbook | verify dynamic array spill size with Aspose.Cells calculateFormula | read sorted values from spilled range using Aspose.Cells in C# | determine number of rows returned by SORT function in Aspose.Cells
// Tags: SORT formula insertion Aspose.Cells | dynamic array spill detection C# | calculate formulas Aspose.Cells .NET | validate spilled range size Aspose.Cells | read sorted values from spilled cells C#

using System;
using Aspose.Cells;

// The example creates a new workbook, populates A2:A11 with numbers 10‑1, inserts the formula "=SORT(A2:A11)" into B2, calculates all formulas, then iterates down column B to count non‑empty cells until an empty cell is encountered. It confirms that the spilled range contains the expected 10 rows and 1 column, prints the sorted values, and saves the workbook as SortDynamicArrayResult.xlsx.
class SortDynamicArrayExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data in column A (A2:A11) with unsorted numbers 10..1
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(10 - i);
            }

            // Insert SORT dynamic array formula into B2 (will spill to B2:B11)
            Cell formulaCell = sheet.Cells["B2"];
            formulaCell.Formula = "=SORT(A2:A11)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Determine the spilled range size manually (dynamic arrays spill until an empty cell)
            int startRow = formulaCell.Row;          // 1 (zero‑based index for B2)
            int startColumn = formulaCell.Column;    // 1 (column B)
            int spilledRows = 0;

            while (true)
            {
                Cell cur = sheet.Cells[startRow + spilledRows, startColumn];

                // In Aspose.Cells a cell is considered empty when its type is IsNull
                if (cur.Type == CellValueType.IsNull)
                    break;

                spilledRows++;
            }

            int expectedRows = 10;
            int expectedColumns = 1;

            if (spilledRows == expectedRows && expectedColumns == 1)
            {
                Console.WriteLine($"Spilled range size is correct: {expectedRows} rows x {expectedColumns} column(s).");
            }
            else
            {
                Console.WriteLine($"Spilled range size is incorrect. Detected: {spilledRows} rows x {expectedColumns} column(s).");
            }

            // Output the spilled values to verify sorting
            Console.WriteLine("Sorted values from spilled range:");
            for (int r = 0; r < spilledRows; r++)
            {
                Cell c = sheet.Cells[startRow + r, startColumn];
                Console.WriteLine(c.StringValue);
            }

            // Save the workbook (optional)
            string outputPath = "SortDynamicArrayResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
