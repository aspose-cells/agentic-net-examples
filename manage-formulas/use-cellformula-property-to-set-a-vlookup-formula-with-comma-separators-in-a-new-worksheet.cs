// Title: Set a VLOOKUP formula with comma separators using Cell.Formula in a newly added worksheet (Aspose.Cells for .NET)
// Common Searches: Aspose.Cells C# assign VLOOKUP formula with commas to a cell | how to use Cell.Formula for VLOOKUP in a newly added worksheet | example of calculating formulas after setting VLOOKUP in Aspose.Cells | save workbook as .xlsx after applying VLOOKUP with Aspose.Cells .NET | populate lookup table and apply VLOOKUP using Aspose.Cells API
// Tags: Cell.Formula VLOOKUP comma separators | add worksheet and fill lookup range Aspose.Cells | calculate workbook formulas Aspose.Cells | save workbook to xlsx Aspose.Cells | C# Aspose.Cells set formula example

using Aspose.Cells;
using System;

// The example creates a new workbook, adds a worksheet named "VLookupSheet", fills columns A and B with key/value data, assigns a VLOOKUP formula to cell C2 using the Cell.Formula property with comma separators, optionally calculates all formulas, and saves the file as VLookupExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Add a new worksheet
        int sheetIndex = workbook.Worksheets.Add();
        var sheet = workbook.Worksheets[sheetIndex];
        sheet.Name = "VLookupSheet";

        // Populate sample data for lookup
        sheet.Cells["A1"].PutValue("Key");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(30);

        // Set a VLOOKUP formula in C2 using comma separators
        // Formula: =VLOOKUP(A2, A:B, 2, FALSE)
        sheet.Cells["C2"].Formula = "VLOOKUP(A2, A:B, 2, FALSE)";

        // Calculate formulas (optional)
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("VLookupExample.xlsx");
    }
}
