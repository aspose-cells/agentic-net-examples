// Title: Set a cell formula that references an external workbook using Aspose.Cells in C#
// Common Searches: how to set a formula that points to another workbook with Aspose.Cells C# | Aspose.Cells example for linking a cell to an external Excel file | C# code to assign external workbook reference in a cell using SetFormula | reference cell A1 from ExternalWorkbook.xlsx in Aspose.Cells | using Aspose.Cells SetFormula to create cross‑workbook formulas in C#
// Tags: Aspose.Cells SetFormula external workbook linking | C# assign cross‑workbook formula | link cell to another Excel workbook Aspose.Cells | external workbook reference formula C# | Aspose.Cells formula referencing external file

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first worksheet, and uses the SetFormula property to place a formula in cell B2 that points to cell A1 on Sheet1 of an external file named ExternalWorkbook.xlsx. The workbook is then saved as Output.xlsx, with exception handling included.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Assign a formula that references an external workbook.
            // Example: reference cell A1 in Sheet1 of ExternalWorkbook.xlsx
            // Formula syntax: ='[ExternalWorkbook.xlsx]Sheet1'!A1
            sheet.Cells["B2"].Formula = "='[ExternalWorkbook.xlsx]Sheet1'!A1";

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
