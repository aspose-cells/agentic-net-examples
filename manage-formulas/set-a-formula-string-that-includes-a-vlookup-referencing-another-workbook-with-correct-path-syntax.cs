// Title: Create a VLOOKUP formula that references an external workbook file path using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set VLOOKUP formula with external workbook reference | How to build Excel VLOOKUP formula that points to another .xlsx file using Aspose.Cells | C# code for VLOOKUP across workbooks with full file path in Aspose.Cells | Aspose.Cells external workbook path syntax in formula string | Set cell formula to lookup value in different workbook using Aspose.Cells .NET
// Tags: VLOOKUP external workbook reference Aspose.Cells | assign formula to cell Aspose.Cells C# | Excel file path syntax in Aspose.Cells formula | cross-workbook lookup using Aspose.Cells | Aspose.Cells .NET external workbook formula

using Aspose.Cells;
using System;

// The example creates a new workbook, builds a VLOOKUP formula string that includes a full path to another Excel file, assigns the formula to cell B2, and saves the workbook as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the external workbook path, sheet name, and lookup range
        string externalWorkbookPath = @"C:\Data\OtherWorkbook.xlsx";
        string externalSheetName = "Sheet1";
        string lookupRange = "$A$1:$B$10";

        // Set a VLOOKUP formula in cell B2 that references the external workbook
        // Formula format: =VLOOKUP(A2,'[C:\Data\OtherWorkbook.xlsx]Sheet1'!$A$1:$B$10,2,FALSE)
        worksheet.Cells["B2"].Formula = $"=VLOOKUP(A2,'[{externalWorkbookPath}]{externalSheetName}'!{lookupRange},2,FALSE)";

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
