// Title: Insert and evaluate a SUM formula for column E rows 2‑100 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, places the formula =SUM(E2:E100) in a chosen cell, forces calculation, and saves the file with Aspose.Cells. | Generate a snippet that sets a SUM range formula on a worksheet, calls CalculateFormula, and outputs the result to an Excel file using Aspose.Cells .NET.
// Common Searches: how to add a SUM(E2:E100) formula to an Excel file with Aspose.Cells C# | Aspose.Cells calculate column total programmatically before saving workbook | C# set formula for range E2 to E100 and evaluate it using Aspose.Cells | using Aspose.Cells to insert and compute a SUM formula in a new workbook
// Tags: Aspose.Cells set cell formula | Aspose.Cells calculate workbook formulas | C# insert SUM range formula | Aspose.Cells save workbook after calculation | Excel column total using Aspose.Cells .NET

using Aspose.Cells;

// // Creates a new workbook, inserts =SUM(E2:E100) into cell A1, evaluates the formula, and saves the workbook as Result.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a formula that sums column E from rows 2 to 100
        // The formula is placed in cell A1 (you can change the target cell as needed)
        Cell targetCell = worksheet.Cells["A1"];
        targetCell.Formula = "=SUM(E2:E100)";

        // Optional: evaluate the formula so the result is stored in the cell
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule)
        workbook.Save("Result.xlsx");
    }
}
