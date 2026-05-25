using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet to the workbook
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet worksheet = workbook.Worksheets[newSheetIndex];

        // Populate a simple lookup table (A2:B4)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Alice");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Bob");
        worksheet.Cells["A4"].PutValue(3);
        worksheet.Cells["B4"].PutValue("Charlie");

        // Value to look up (placed in D1)
        worksheet.Cells["D1"].PutValue(2);

        // Set VLOOKUP formula in E1 using comma as argument separator
        // Syntax: =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
        worksheet.Cells["E1"].Formula = "=VLOOKUP(D1, A2:B4, 2, FALSE)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("VlookupDemo.xlsx");
    }
}