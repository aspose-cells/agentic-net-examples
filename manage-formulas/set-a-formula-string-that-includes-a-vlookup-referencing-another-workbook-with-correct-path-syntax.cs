// Title: Set a VLOOKUP formula with an external workbook full file path using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set VLOOKUP formula that points to another workbook using full file path | Create external workbook reference in VLOOKUP formula with Aspose.Cells .NET | C# Aspose.Cells formula syntax for cross‑workbook VLOOKUP | How to include full path in VLOOKUP formula using Aspose.Cells API | Set cell formula to VLOOKUP external workbook in Aspose.Cells C# example
// Tags: VLOOKUP external workbook Aspose.Cells | formula absolute path syntax Aspose.Cells | cross workbook lookup C# Aspose.Cells | set cell formula Aspose.Cells | external sheet reference Aspose.Cells

using Aspose.Cells;

// Demonstrates how to build a VLOOKUP formula that references another workbook via a full file path, assign it to a cell, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the external workbook path and sheet
        string externalPath = @"C:\Data\[OtherWorkbook.xlsx]Sheet1";

        // Build the VLOOKUP formula referencing the external workbook
        // Syntax: =VLOOKUP(lookup_value, 'full_path\[workbook.xlsx]SheetName'!range, col_index, [range_lookup])
        string formula = $"=VLOOKUP(A2,'{externalPath}'!$A$1:$B$10,2,FALSE)";

        // Set the formula in cell C2
        sheet.Cells["C2"].Formula = formula;

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}
