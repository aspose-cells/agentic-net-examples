using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.DigitalSignatures;
using AsposeRange = Aspose.Cells.Range;

class EnumeratorScenarios
{
    static void Main()
    {
        // ------------------------------------------------------------
        // Scenario 1: Enumerate all cells in a worksheet.
        // ------------------------------------------------------------
        Workbook wb1 = new Workbook();
        Worksheet ws1 = wb1.Worksheets[0];
        ws1.Cells["A1"].PutValue("Name");
        ws1.Cells["B1"].PutValue("Age");
        ws1.Cells["A2"].PutValue("John");
        ws1.Cells["B2"].PutValue(30);

        IEnumerator cellEnum = ws1.Cells.GetEnumerator();
        Console.WriteLine("All cells with values:");
        while (cellEnum.MoveNext())
        {
            Cell c = (Cell)cellEnum.Current;
            if (c.Value != null)
                Console.WriteLine($"{c.Name}: {c.Value}");
        }

        // ------------------------------------------------------------
        // Scenario 2: Enumerate rows with a synchronized enumerator.
        // ------------------------------------------------------------
        Workbook wb2 = new Workbook();
        Worksheet ws2 = wb2.Worksheets[0];
        for (int i = 0; i < 5; i++)
            ws2.Cells[i, 0].PutValue($"Row {i}");

        IEnumerator rowEnum = ws2.Cells.Rows.GetEnumerator();
        Console.WriteLine("\nRows (synchronized):");
        while (rowEnum.MoveNext())
        {
            Row r = (Row)rowEnum.Current;
            Console.WriteLine($"Row {r.Index}: {r[0].StringValue}");

            // Dynamically insert a new row after processing row index 2.
            if (r.Index == 2)
                ws2.Cells[10, 0].PutValue("Inserted Row");
        }

        // ------------------------------------------------------------
        // Scenario 3: Enumerate cells within a specific range.
        // ------------------------------------------------------------
        Workbook wb3 = new Workbook();
        Worksheet ws3 = wb3.Worksheets[0];
        ws3.Cells["B2"].PutValue("Apple");
        ws3.Cells["C2"].PutValue("Banana");
        ws3.Cells["B3"].PutValue("Cherry");
        ws3.Cells["C3"].PutValue("Date");

        AsposeRange rng = ws3.Cells.CreateRange("B2:C3");
        IEnumerator rangeEnum = rng.GetEnumerator();
        Console.WriteLine("\nCells in range B2:C3:");
        while (rangeEnum.MoveNext())
        {
            Cell c = (Cell)rangeEnum.Current;
            Console.WriteLine($"{c.Name}: {c.Value}");
        }

        // ------------------------------------------------------------
        // Scenario 4: Enumerate pivot fields in a pivot table.
        // ------------------------------------------------------------
        Workbook wb4 = new Workbook();
        Worksheet ws4 = wb4.Worksheets[0];
        ws4.Cells["A1"].PutValue("Product");
        ws4.Cells["B1"].PutValue("Sales");
        ws4.Cells["A2"].PutValue("Laptop");
        ws4.Cells["B2"].PutValue(1200);
        ws4.Cells["A3"].PutValue("Phone");
        ws4.Cells["B3"].PutValue(800);
        ws4.Cells["A4"].PutValue("Tablet");
        ws4.Cells["B4"].PutValue(600);

        int ptIdx = ws4.PivotTables.Add("A1:B4", "E5", "SalesPivot");
        PivotTable pt = ws4.PivotTables[ptIdx];
        pt.AddFieldToArea(PivotFieldType.Row, 0);
        pt.AddFieldToArea(PivotFieldType.Data, 1);
        pt.RefreshData();
        pt.CalculateData();

        IEnumerator pivotFieldEnum = pt.RowFields.GetEnumerator();
        Console.WriteLine("\nPivot row fields:");
        while (pivotFieldEnum.MoveNext())
        {
            PivotField pf = (PivotField)pivotFieldEnum.Current;
            Console.WriteLine(pf.Name);
        }

        // ------------------------------------------------------------
        // Save a sample workbook to demonstrate proper lifecycle handling.
        // ------------------------------------------------------------
        wb1.Save("EnumeratorDemo.xlsx");
    }
}