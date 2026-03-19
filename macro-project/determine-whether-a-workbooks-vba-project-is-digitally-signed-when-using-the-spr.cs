using System;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Load the workbook saved in SpreadsheetML format (e.g., .xlsx or .xlsm)
        Workbook workbook = new Workbook("input.xlsx");

        // Verify that the workbook contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Determine whether the VBA project is digitally signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project is signed: " + isSigned);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}