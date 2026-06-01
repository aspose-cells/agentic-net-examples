using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // ---------- Populate initial data in column A ----------
            cells["A1"].PutValue("Item1");
            cells["A2"].PutValue("Item2");
            cells["A3"].PutValue("Item3");

            // ---------- Set a dynamic array formula in B1 ----------
            // The formula creates a sequence whose length equals the number of non‑empty cells in column A.
            // SEQUENCE(COUNTA(A:A)) will spill into B1, B2, B3, … automatically.
            cells["B1"].SetDynamicArrayFormula("=SEQUENCE(COUNTA(A:A))", new FormulaParseOptions(), true);

            // Calculate and refresh the dynamic array so the spill range is materialized.
            wb.CalculateFormula();
            wb.RefreshDynamicArrayFormulas(true);

            // ---------- Define a named range that points to the spilled range ----------
            // The “#” after a cell reference returns the entire spill range of that cell.
            int nameIdx = wb.Worksheets.Names.Add("MyDynamicRange");
            wb.Worksheets.Names[nameIdx].RefersTo = "=Sheet1!$B$1#";

            // ---------- Read values via the named range ----------
            Name dynName = wb.Worksheets.Names["MyDynamicRange"];
            AsposeRange dynRange = dynName.GetRange(); // current spilled range

            Console.WriteLine("Initial dynamic range values:");
            for (int i = 0; i < dynRange.RowCount; i++)
            {
                Console.WriteLine(dynRange[i, 0].Value);
            }

            // ---------- Add more data to column A ----------
            cells["A4"].PutValue("Item4");
            cells["A5"].PutValue("Item5");

            // Re‑calculate and refresh the dynamic array so the spill expands.
            wb.CalculateFormula();
            wb.RefreshDynamicArrayFormulas(true);

            // ---------- Read updated values ----------
            dynRange = dynName.GetRange(); // refreshed spilled range
            Console.WriteLine("\nAfter adding more items:");
            for (int i = 0; i < dynRange.RowCount; i++)
            {
                Console.WriteLine(dynRange[i, 0].Value);
            }

            // ---------- Save the workbook ----------
            string outputPath = "DynamicNamedRangeDemo.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}