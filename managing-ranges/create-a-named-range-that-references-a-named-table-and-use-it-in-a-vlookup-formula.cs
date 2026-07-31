// Title: Aspose.Cells for .NET (C#): Create a Named Range from a ListObject Table and Use It in a VLOOKUP Formula
// Description: This C# example shows how to build a workbook, add a ListObject table, define a workbook‑level named range that points to the table, apply the named range in a VLOOKUP formula, calculate the result, and save the file.
// Keywords: Aspose.Cells | C# | named range | ListObject | Excel table | VLOOKUP formula | reference table in formula | Workbook.CalculateFormula | Aspose.Cells example | GitHub demo | Excel automation
// Common Searches: Aspose.Cells create named range from table C# | C# VLOOKUP using named range Aspose.Cells | How to reference a ListObject in a VLOOKUP with Aspose.Cells | Aspose.Cells named range example for formulas | C# Excel automation VLOOKUP named range
// Developer Intent: Define a named range that references a ListObject table and use that range in a VLOOKUP formula with Aspose.Cells for .NET.
// Use Cases: Expose a product‑price table via a named range so multiple worksheets can perform price lookups without duplicating data. | Maintain a single source of truth: updating the ListObject automatically refreshes all VLOOKUP results that use the named range. | Create reusable formula components for templates that require dynamic table references across workbooks.
// AI Prompts: Generate C# code using Aspose.Cells to add a ListObject table, create a workbook‑level named range that points to the table, and write a VLOOKUP formula referencing that named range. | Show how to set the RefersTo property of a workbook name to a ListObject name and evaluate the VLOOKUP result with Workbook.CalculateFormula(). | Provide a sample that reuses a named range referencing a table in VLOOKUP formulas on different worksheets in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

namespace AsposeCellsNamedRangeVLookupDemo
{
    // This C# example shows how to build a workbook, add a ListObject table, define a workbook‑level named range that points to the table, apply the named range in a VLOOKUP formula, calculate the result, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data for a table (Product, Price)
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.50);

                // ---------- Create a table (ListObject) ----------
                // The table will cover A1:B4 (zero‑based indices)
                int firstRow = 0;
                int firstCol = 0;
                int totalRows = 4;
                int totalCols = 2;

                // Add returns the index of the newly created ListObject
                int tableIndex = sheet.ListObjects.Add(
                    firstRow,
                    firstCol,
                    firstRow + totalRows - 1,
                    firstCol + totalCols - 1,
                    true);

                // Retrieve the ListObject and set its name
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "ProductsTable"; // Use DisplayName (or Name if supported)

                // ---------- Create a named range that references the table ----------
                // Add a new name to the workbook's name collection
                int nameIndex = workbook.Worksheets.Names.Add("ProductsRange");
                // Set the RefersTo property to the table name (preceded by '=')
                workbook.Worksheets.Names[nameIndex].RefersTo = "=ProductsTable";

                // ---------- Use the named range in a VLOOKUP formula ----------
                // Example: lookup the price of "Banana"
                sheet.Cells["D1"].PutValue("Lookup Product");
                sheet.Cells["E1"].PutValue("Price (VLOOKUP)");
                sheet.Cells["D2"].PutValue("Banana");
                // VLOOKUP(lookup_value, named_range, column_index, FALSE)
                sheet.Cells["E2"].Formula = "=VLOOKUP(D2, ProductsRange, 2, FALSE)";

                // Calculate formulas so that the VLOOKUP result is evaluated
                workbook.CalculateFormula();

                // ---------- Save the workbook ----------
                string outputPath = "NamedRangeVLookupDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
