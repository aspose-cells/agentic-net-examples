// Title: Treat blank cells as zero when summing a mixed range with Aspose.Cells CalculationOptions in C#
// Common Searches: Aspose.Cells calculate sum ignoring blank cells C# | set CalculationOptions treat empty cells as zero Aspose.Cells .NET | evaluate Excel SUM formula with blanks using Aspose.Cells API | C# workbook.CalculateFormula treat empty cells as zero | how to configure Aspose.Cells to treat empty cells as zero in formula evaluation
// Tags: Aspose.Cells CalculationOptions blank cells zero | C# sum formula with empty cells Aspose.Cells | evaluate Excel SUM range Aspose.Cells .NET | set calculation options for blank cells Aspose.Cells | formula calculation workbook.CalculateFormula Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills A1, A2, and B2 with numbers while leaving B1 empty, configures CalculationOptions to treat empty cells as zero, assigns a SUM(A1:B2) formula to C1, calculates the workbook, and prints the resulting total.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some cells; leave others empty
            cells["A1"].PutValue(10);   // numeric value
            cells["A2"].PutValue(20);   // numeric value
            // B1 remains empty
            cells["B2"].PutValue(30);   // numeric value

            // Place the formula in a cell
            cells["C1"].Formula = "=SUM(A1:B2)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the calculated result
            double sumResult = cells["C1"].DoubleValue;

            // Display the result
            Console.WriteLine("Sum result: " + sumResult);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
