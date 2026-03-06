using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNonSequentialRanges
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook (lifecycle create) ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in two separate blocks
            // Block 1: A1:A3
            cells["A1"].PutValue("Item");
            cells["A2"].PutValue("Apple");
            cells["A3"].PutValue("Banana");

            // Block 2: C1:C3
            cells["C1"].PutValue("Price");
            cells["C2"].PutValue(1.20);
            cells["C3"].PutValue(0.80);

            // ---------- Create non‑sequential ranges ----------
            // First range (A1:A3)
            AsposeRange range1 = cells.CreateRange("A1", "A3");
            range1.Name = "ItemsRange";

            // Second range (C1:C3)
            AsposeRange range2 = cells.CreateRange("C1", "C3");
            range2.Name = "PricesRange";

            // Add both ranges to the worksheet's Cells collection.
            // This does not merge them; it simply registers the ranges.
            cells.AddRange(range1);
            cells.AddRange(range2);

            // ---------- Export each range to a DataTable ----------
            DataTable dtItems = range1.ExportDataTable();
            DataTable dtPrices = range2.ExportDataTable();

            // Simple console output to verify the exported data
            Console.WriteLine("Items Range:");
            foreach (DataRow row in dtItems.Rows)
                Console.WriteLine($"  {row[0]}");

            Console.WriteLine("\nPrices Range:");
            foreach (DataRow row in dtPrices.Rows)
                Console.WriteLine($"  {row[0]}");

            // ---------- Save the workbook (lifecycle save) ----------
            // The file will contain the original data and the named ranges.
            workbook.Save("NonSequentialRanges.xlsx");
        }
    }
}