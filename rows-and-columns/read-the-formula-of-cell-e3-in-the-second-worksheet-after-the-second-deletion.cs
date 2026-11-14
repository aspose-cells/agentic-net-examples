// Title: Read the formula in cell E3 of the second worksheet after deleting rows in the first worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read formula from E3 on second sheet after deleting rows in first sheet | how to retrieve a cell formula after row deletions using Aspose.Cells in .NET | C# Aspose.Cells get formula from a specific worksheet after modifying another worksheet | read cell formula from another worksheet after removing rows with Aspose.Cells
// Tags: read cell formula Aspose.Cells C# | delete rows first worksheet Aspose.Cells | access second worksheet cell E3 Aspose.Cells | retrieve formula after row deletion Aspose.Cells | Aspose.Cells workbook row removal impact formula

using Aspose.Cells;
using System;
using System.IO;

// The example creates or loads an Excel file, deletes the first two rows of the first worksheet, then accesses the second worksheet and reads the formula stored in cell E3 (e.g., "=SUM(A1:A2)"), outputting the formula to the console using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Ensure the input file exists; create a minimal workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                // Add a second worksheet with a sample formula in E3.
                var ws2 = tempWb.Worksheets[tempWb.Worksheets.Add()];
                ws2.Name = "SecondSheet";
                ws2.Cells["E3"].Formula = "=SUM(A1:A2)";
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // First deletion: delete the first row of the first worksheet.
            Worksheet wsFirst = workbook.Worksheets[0];
            wsFirst.Cells.DeleteRows(0, 1); // startRow = 0, totalRows = 1

            // Second deletion: delete the new first row (originally the second row).
            wsFirst.Cells.DeleteRows(0, 1);

            // Access the second worksheet (index 1).
            if (workbook.Worksheets.Count > 1)
            {
                Worksheet wsSecond = workbook.Worksheets[1];

                // Retrieve the formula from cell E3.
                Cell targetCell = wsSecond.Cells["E3"];
                string formula = targetCell.Formula;

                // Output the formula.
                Console.WriteLine("Formula in E3 of the second worksheet: " + formula);
            }
            else
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
