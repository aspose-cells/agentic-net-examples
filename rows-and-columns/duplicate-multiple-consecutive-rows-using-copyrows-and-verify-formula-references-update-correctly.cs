// Title: Copy multiple consecutive rows with Aspose.Cells CopyRows and verify that formulas adjust to the new cell references in C#
// AI Prompts: Invoke Cells.CopyRows to duplicate rows 0‑2 into rows 3‑5 and display the formula in the copied B5 cell. | After the copy operation, call Workbook.CalculateFormula and output the evaluated value of cell B5. | Implement try‑catch blocks to log any errors that occur during row copying, formula calculation, or workbook saving. | Save the modified workbook as an .xlsx file and confirm that the file was created successfully.
// Common Searches: Aspose.Cells C# copy rows and keep relative formula references | How does Cells.CopyRows adjust formulas when duplicating rows in .NET | Example of copying rows with formulas using Aspose.Cells and verifying results | Recalculating workbook after using CopyRows in Aspose.Cells C# | CopyRows method preserving formula links in Excel file with Aspose.Cells
// Tags: CopyRows with relative formula adjustment | duplicate rows preserving Excel formulas Aspose.Cells | calculate workbook after row copy C# | save workbook as .xlsx after copying rows | exception handling for Cells.CopyRows operation

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, fills rows 0‑2 with values and a formula referencing A1, uses Cells.CopyRows to duplicate those rows to positions 3‑5, checks that the copied formula in B5 now points to the new A4 cell, recalculates the workbook to obtain the formula result, prints the outcomes, and saves the file as CopyRowsFormulaUpdateDemo.xlsx while handling potential exceptions.
    public class CopyRowsFormulaUpdateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate source rows (rows 0-2)
                // Row 0: simple value
                cells[0, 0].PutValue(10);               // A1 = 10

                // Row 1: formula that references the cell above (A1)
                cells[1, 1].Formula = "=A1*2";          // B2 = A1*2

                // Row 2: another simple value
                cells[2, 0].PutValue(30);               // A3 = 30

                // Verify original formula
                Console.WriteLine("Original formula in B2: " + cells[1, 1].Formula);

                // Duplicate the three rows (0,1,2) starting at destination row index 3
                // This will copy rows 0‑2 to rows 3‑5
                cells.CopyRows(cells, 0, 3, 3);

                // After copying, the formula in the copied row (row index 4, column 1) should be adjusted
                // It should now reference the copied A4 cell (which contains the value 10)
                Console.WriteLine("Copied formula in B5: " + cells[4, 1].Formula);

                // Optional: calculate to see the result of the copied formula
                workbook.CalculateFormula();
                Console.WriteLine("Value of B5 after calculation: " + cells[4, 1].Value);

                // Save the workbook
                workbook.Save("CopyRowsFormulaUpdateDemo.xlsx");
                Console.WriteLine("Workbook saved as CopyRowsFormulaUpdateDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyRowsFormulaUpdateDemo.Run();
        }
    }
}
