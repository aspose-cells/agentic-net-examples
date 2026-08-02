using System;
using Aspose.Cells;

class DynamicNamedRangeDemo
{
    static void Main()
    {
        // 1. Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // 2. Populate initial data in column A (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1); // values 1..5
        }

        // 3. Define a dynamic named range "MyData" that expands with non‑empty cells in column A
        //    Formula uses OFFSET together with COUNTA to calculate the height dynamically.
        int nameIndex = workbook.Worksheets.Names.Add("MyData");
        Name myDataName = workbook.Worksheets.Names[nameIndex];
        myDataName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

        // 4. Set a dynamic array formula that consumes the named range.
        //    Example: sort the data automatically.
        Cell formulaCell = cells["B1"];
        formulaCell.SetDynamicArrayFormula("=SORT(MyData)", new FormulaParseOptions(), true);

        // 5. Calculate formulas and refresh dynamic array spill range
        workbook.CalculateFormula();
        workbook.RefreshDynamicArrayFormulas(true);

        // 6. Append new data rows to column A (A6 and A7)
        cells[5, 0].PutValue(6);
        cells[6, 0].PutValue(7);

        // 7. Refresh dynamic array formulas again so the spill range expands automatically
        workbook.RefreshDynamicArrayFormulas(true);
        workbook.CalculateFormula();

        // 8. Save the workbook
        workbook.Save("DynamicNamedRangeDemo.xlsx");
    }
}