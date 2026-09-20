// Title: Add a VLOOKUP formula with comma separators to a new worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook, add a worksheet named "VLookupSheet", populate A1:B5 with a key‑value table, assign a VLOOKUP expression with commas to cell C2 via the Formula property, and save the file as an .xlsx workbook. | Generate an Excel file in C# that demonstrates setting a VLOOKUP formula (using comma argument separators) on a newly created worksheet using Aspose.Cells and then persisting the workbook.
// Common Searches: Aspose.Cells C# set VLOOKUP formula with commas in a specific cell | example of using Cell.Formula to insert VLOOKUP in a new worksheet with Aspose.Cells | how to write VLOOKUP(lookup_value, table_array, col_index, FALSE) using Aspose.Cells .NET
// Tags: Aspose.Cells set cell VLOOKUP formula | C# insert VLOOKUP with comma separators | create worksheet and apply lookup formula Aspose.Cells | save workbook as XLSX after formula assignment

using Aspose.Cells;

// Demonstrates creating a new workbook, adding a worksheet named 'VLookupSheet', filling a lookup table in A1:B5, setting a VLOOKUP formula with comma separators in cell C2 via the Cell.Formula property, and saving the workbook as VLookupExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet and give it a name
        int sheetIndex = workbook.Worksheets.Add();
        Worksheet sheet = workbook.Worksheets[sheetIndex];
        sheet.Name = "VLookupSheet";

        // Populate sample lookup table in A1:B5
        sheet.Cells["A1"].PutValue("Key");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue("Date");
        sheet.Cells["B5"].PutValue(40);

        // Set a VLOOKUP formula in C2 using comma separators
        // Syntax: VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
        sheet.Cells["C2"].Formula = "VLOOKUP(\"Banana\",A1:B5,2,FALSE)";

        // Save the workbook to a file
        workbook.Save("VLookupExample.xlsx");
    }
}
