using System;
using Aspose.Cells;

class DynamicNamedRangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Data";

        // Populate the first row (header) with initial columns
        ws.Cells["A1"].PutValue("Header1");
        ws.Cells["B1"].PutValue("Header2");
        ws.Cells["C1"].PutValue("Header3");

        // Add some sample data under the headers
        ws.Cells["A2"].PutValue(10);
        ws.Cells["B2"].PutValue(20);
        ws.Cells["C2"].PutValue(30);

        // Create a dynamic named range that expands horizontally as columns are added.
        // The formula uses OFFSET together with COUNTA to count non‑empty cells in row 1.
        // =OFFSET(Data!$A$1,0,0,1,COUNTA(Data!$1:$1))
        int nameIdx = wb.Worksheets.Names.Add("Headers");
        Name dynamicName = wb.Worksheets.Names[nameIdx];
        dynamicName.RefersTo = "=OFFSET(Data!$A$1,0,0,1,COUNTA(Data!$1:$1))";

        // Use the named range in a formula to verify it works (counts columns)
        ws.Cells["E1"].Formula = "=COLUMNS(Headers)";

        // Calculate formulas so that E1 shows the current column count
        wb.CalculateFormula();

        // Insert a new column to the right of the existing data (index 3 = column D)
        ws.Cells.InsertColumn(3);
        ws.Cells["D1"].PutValue("Header4");
        ws.Cells["D2"].PutValue(40);

        // Recalculate to let the dynamic named range pick up the new column
        wb.CalculateFormula();

        // Save the workbook
        wb.Save("DynamicNamedRange.xlsx");
    }
}