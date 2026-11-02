// Title: C# – Remove a cell from the Aspose.Cells Watch Window after CalculateFormula execution
// Common Searches: Aspose.Cells C# clear watch window entry after CalculateFormula | how to programmatically remove a cell from watch window in Aspose.Cells .NET | remove watched cell after formula evaluation using Aspose.Cells API | C# Aspose.Cells watch window cleanup after workbook.CalculateFormula
// Tags: Aspose.Cells watch window cleanup | remove cell from watch list .NET | CalculateFormula watch window management | C# Aspose.Cells formula evaluation cleanup | watch window API Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, assigns a SUM formula to cell A1 on the first worksheet, evaluates all formulas with CalculateFormula, and saves the workbook as 'output.xlsx', handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place a formula in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=SUM(10,20,30)";

            // Evaluate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
