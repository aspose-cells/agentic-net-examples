using System;
using Aspose.Cells;

namespace AsposeCellsHeaderAutoFitDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Populate header rows ----------
            // Row 0 - Header titles
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Description");
            cells["C1"].PutValue("Price");

            // Row 1 - Sub‑header (e.g., units)
            cells["A2"].PutValue("ID");
            cells["B2"].PutValue("Details");
            cells["C2"].PutValue("USD");

            // ---------- Populate data rows ----------
            // Row 2
            cells["A3"].PutValue("001");
            cells["B3"].PutValue("A short description");
            cells["C3"].PutValue(9.99);

            // Row 3 - longer text to demonstrate row height change
            cells["A4"].PutValue("002");
            cells["B4"].PutValue("This is a much longer description that should cause the row to expand when auto‑fitted.");
            cells["C4"].PutValue(19.99);

            // Row 4
            cells["A5"].PutValue("003");
            cells["B5"].PutValue("Another description");
            cells["C5"].PutValue(5.50);

            // ---------- Copy data rows to a new location ----------
            // Copy rows 2‑4 (indices 2‑4) to rows 5‑7 (starting at index 5)
            // rowNumber = 3 because we copy three rows
            cells.CopyRows(cells, 2, 5, 3);

            // ---------- AutoFit header rows only ----------
            // This keeps header height consistent regardless of data rows
            sheet.AutoFitRows(0, 1);

            // ---------- AutoFit the rest of the rows ----------
            // Adjust heights for data rows (original and copied)
            sheet.AutoFitRows(2, 7);

            // Save the workbook to a file
            workbook.Save("HeaderAutoFitDemo.xlsx");
        }
    }
}