using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook (lifecycle: create) ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Define smart markers in the template ----------
            // Row 1 contains the smart markers that will be replaced by data source values.
            // The last column uses a smart marker named "Formula" which holds a formula string.
            cells["A1"].PutValue("&=ProductID");
            cells["B1"].PutValue("&=Quantity");
            cells["C1"].PutValue("&=Price");
            cells["D1"].PutValue("&=Formula"); // This will be set as a formula in the cell.

            // ---------- Prepare the data source ----------
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("ProductID", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(double));
            dt.Columns.Add("Formula", typeof(string)); // Holds the formula string for each row.

            // Populate rows. The formula references the same row's Quantity (B) and Price (C).
            for (int i = 2; i <= 5; i++) // Create 4 data rows.
            {
                DataRow row = dt.NewRow();
                row["ProductID"] = $"P{i - 1}";
                row["Quantity"] = i * 10;               // Example quantity.
                row["Price"] = i * 1.5;                  // Example price.
                // Build a formula that multiplies Quantity (B) and Price (C) of the current row.
                row["Formula"] = $"=B{i}*C{i}";
                dt.Rows.Add(row);
            }

            // ---------- Set up the designer ----------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            // Enable automatic calculation after smart markers are processed.
            designer.CalculateFormula = true;

            // ---------- Process the smart markers ----------
            designer.Process();

            // ---------- Save the result (lifecycle: save) ----------
            workbook.Save("SmartMarker_With_Formula.xlsx");

            Console.WriteLine("Workbook generated with dynamic formulas evaluated.");
        }
    }
}