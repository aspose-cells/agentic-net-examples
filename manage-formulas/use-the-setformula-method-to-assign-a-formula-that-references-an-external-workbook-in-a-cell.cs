// Title: Set a cell formula that references an external workbook with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set formula to external workbook cell | How to link a cell to another Excel file using SetFormula in Aspose.Cells | C# example of external workbook reference in Excel formula with Aspose.Cells | Assign external workbook reference to a cell using Aspose.Cells API
// Tags: setformula external workbook reference c# | aspocells link cell to another workbook | c# assign external workbook formula | excel external workbook formula aspocells | save workbook after setting external formula c#

using System;
using Aspose.Cells;

// Demonstrates using Aspose.Cells in C# to set cell A1's formula to ='[ExternalWorkbook.xlsx]Sheet1'!B2, referencing another workbook, and then saving the result.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula in cell A1 that references an external workbook
        // Example: reference cell B2 from Sheet1 of ExternalWorkbook.xlsx located in the same folder
        string externalFormula = "='[ExternalWorkbook.xlsx]Sheet1'!B2";
        sheet.Cells["A1"].Formula = externalFormula; // using the Formula property (equivalent to SetFormula)

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
