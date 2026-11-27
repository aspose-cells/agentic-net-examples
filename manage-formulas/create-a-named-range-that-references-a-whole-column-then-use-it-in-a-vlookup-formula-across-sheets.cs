// Title: How to define a whole‑column named range and use it in a VLOOKUP formula across worksheets with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create named range that refers to an entire column | C# VLOOKUP formula using a named range across different worksheets in Aspose.Cells | reference a table named range in VLOOKUP with Aspose.Cells for .NET | calculate formulas after setting VLOOKUP with named range in Aspose.Cells | save workbook after using VLOOKUP with a named range in Aspose.Cells C#
// Tags: create whole column named range Aspose.Cells C# | VLOOKUP with named range across worksheets Aspose.Cells | define table named range for lookup Aspose.Cells | programmatic formula calculation Aspose.Cells .NET | save workbook after VLOOKUP Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, defines a whole‑column named range (KeyColumn) and a two‑column table named range (KeyTable) on the Data sheet, adds a Lookup sheet, inserts a lookup key, sets a VLOOKUP formula that references the KeyTable named range, calculates the formulas, and saves the file as NamedRangeVlookup.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (source data)
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Data";

            // Fill sample data: column B = key, column C = value
            srcSheet.Cells["B1"].PutValue("Key");
            srcSheet.Cells["C1"].PutValue("Value");
            srcSheet.Cells["B2"].PutValue("A");
            srcSheet.Cells["C2"].PutValue(100);
            srcSheet.Cells["B3"].PutValue("B");
            srcSheet.Cells["C3"].PutValue(200);
            srcSheet.Cells["B4"].PutValue("C");
            srcSheet.Cells["C4"].PutValue(300);

            // Create a named range that references the whole column B in the Data sheet
            int keyColIdx = workbook.Worksheets.Names.Add("KeyColumn");
            Name keyColName = workbook.Worksheets.Names[keyColIdx];
            keyColName.RefersTo = "Data!B:B";

            // Create a named range that covers columns B and C (key + return column)
            int keyTableIdx = workbook.Worksheets.Names.Add("KeyTable");
            Name keyTableName = workbook.Worksheets.Names[keyTableIdx];
            keyTableName.RefersTo = "Data!B:C";

            // Add a second worksheet where the VLOOKUP will be used
            int lookupSheetIdx = workbook.Worksheets.Add();
            Worksheet lookupSheet = workbook.Worksheets[lookupSheetIdx];
            lookupSheet.Name = "Lookup";

            // Place a lookup key in cell A2
            lookupSheet.Cells["A2"].PutValue("B");

            // Set VLOOKUP formula in B2 using the named range KeyTable
            lookupSheet.Cells["B2"].Formula = "=VLOOKUP(A2,KeyTable,2,FALSE)";

            // Calculate formulas so the result is stored in the cell
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "NamedRangeVlookup.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
