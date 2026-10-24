// Title: Setting AutomaticExceptTables mode and adding a ListObject table to test formula recalculation with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set calculation mode to AutomaticExceptTables and recalculate formulas | Add a ListObject table in Aspose.Cells and verify SUM formula updates automatically | How to enable AutomaticExceptTables mode in Aspose.Cells .NET and observe totals row behavior | Recalculate workbook after inserting a table with a totals row using Aspose.Cells C# | Example of using workbook.CalculateFormula after adding a ListObject in Aspose.Cells
// Tags: AutomaticExceptTables mode Aspose.Cells | create ListObject Aspose.Cells C# | recalculate formulas after table insertion Aspose.Cells | Excel totals row with Aspose.Cells ListObject | save workbook as .xlsx using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a new workbook, optionally sets the calculation mode to AutomaticExceptTables, fills cells A1:B5 with sample data, places a SUM formula in D1, forces an initial calculation, then adds a ListObject table over the data range with a totals row that sums column B. After inserting the table, the workbook is recalculated to show the updated sum in D1 and the table total, and finally the workbook is saved as CalculationModeExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Optional: set calculation mode if supported by the library version
            // workbook.Settings.CalcMode = CalculationMode.AutomaticExceptTables;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in A1:B5
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);          // Column A: 1,2,3,4,5
                sheet.Cells[i, 1].PutValue((i + 1) * 10);   // Column B: 10,20,30,40,50
            }

            // Add a formula outside the table that sums column B
            sheet.Cells["D1"].Formula = "=SUM(B1:B5)";

            // Force calculation to obtain the initial result
            workbook.CalculateFormula();

            Console.WriteLine("Initial sum (D1): " + sheet.Cells["D1"].Value);

            // Define the range for the new table (A1:B5)
            int firstRow = 0;
            int firstCol = 0;
            int totalRows = 5;   // number of data rows
            int totalCols = 2;   // number of columns

            // Add the table (ListObject) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstCol,
                firstRow + totalRows - 1,
                firstCol + totalCols - 1,
                true); // has headers

            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Show a total row and set the total for column B to SUM
            table.ShowTotals = true;
            table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum; // Column B (zero‑based index 1)

            // Recalculate formulas after adding the table
            workbook.CalculateFormula();

            Console.WriteLine("After adding table, sum (D1): " + sheet.Cells["D1"].Value);
            // The total row for column B is placed just below the data (row index = totalRows)
            Console.WriteLine("Table total (B total row): " + sheet.Cells[totalRows, 1].Value);

            // Save the workbook
            string outputPath = "CalculationModeExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
