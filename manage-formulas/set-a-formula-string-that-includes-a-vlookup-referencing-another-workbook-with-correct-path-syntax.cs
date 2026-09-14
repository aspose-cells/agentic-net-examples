// Title: Add a VLOOKUP formula with an external workbook reference using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, builds a VLOOKUP formula that points to a range in another Excel file using the correct external reference syntax, and writes the formula to a cell. | Show how to format the external workbook path and sheet name in a VLOOKUP string for Aspose.Cells, then assign it to a cell programmatically. | Provide a step‑by‑step example of constructing and applying a cross‑file VLOOKUP formula in Aspose.Cells, including saving the resulting workbook.
// Common Searches: how to set a VLOOKUP formula that references another Excel workbook in Aspose.Cells C# | Aspose.Cells external workbook VLOOKUP syntax example | C# create VLOOKUP with external file path using Aspose.Cells | building cross‑workbook lookup formula in .NET Aspose.Cells
// Tags: Aspose.Cells set external VLOOKUP formula | C# construct Excel VLOOKUP with file reference | cross‑workbook lookup using Aspose.Cells | external range syntax in Aspose.Cells formulas | assign formula to cell Aspose.Cells .NET

using System;
using Aspose.Cells;

// Demonstrates creating a new workbook with Aspose.Cells, constructing a VLOOKUP formula that references a range in an external Excel file using proper path syntax, assigning the formula to cell B2, and saving the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the external workbook path and sheet name
        string externalWorkbookPath = @"C:\Data\Source.xlsx";
        string externalSheetName = "Data";

        // Define the lookup range inside the external sheet
        string externalRange = "$A$1:$B$10";

        // Build the VLOOKUP formula that references the external workbook
        // Excel syntax: =VLOOKUP(A2,'C:\Data\[Source.xlsx]Data'!$A$1:$B$10,2,FALSE)
        string formula = $"=VLOOKUP(A2,'{externalWorkbookPath.Replace("\\", "/")}[{System.IO.Path.GetFileName(externalWorkbookPath)}]{externalSheetName}'!{externalRange},2,FALSE)";

        // Set the formula in cell B2 (adjust as needed)
        sheet.Cells["B2"].Formula = formula;

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
