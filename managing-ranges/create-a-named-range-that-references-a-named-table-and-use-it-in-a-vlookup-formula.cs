// Title: Aspose.Cells .NET – Define a Named Range from a ListObject and Apply VLOOKUP (C#)
// Description: This C# example demonstrates how to create a workbook, add a ListObject (table) named "ProductsTable", define a workbook‑level named range "ProductRange" that points to the table, insert a VLOOKUP formula using the named range to retrieve the price of "Orange", calculate the formula, output the result, and save the file as an .xlsx document.
// Keywords: Aspose.Cells | C# | .NET | named range | ListObject | table reference | VLOOKUP formula | formula calculation | workbook save | Excel automation
// Common Searches: Aspose.Cells create named range from table | C# VLOOKUP using named range Aspose.Cells | How to reference ListObject in a name Aspose.Cells | Calculate formulas after VLOOKUP Aspose.Cells .NET | Sample code for named range and VLOOKUP in Aspose.Cells
// Developer Intent: Define a named range that points to a table and use it in a VLOOKUP formula with Aspose.Cells for .NET.
// Use Cases: Create a ListObject, assign a display name, and expose it through a workbook‑level name for cleaner formulas. | Apply VLOOKUP with the named range as the table array to look up values without hard‑coding cell addresses. | Force formula evaluation programmatically and persist the result in the saved workbook.
// AI Prompts: Write C# code using Aspose.Cells to add a table, create a named range that references the table, and set a VLOOKUP formula that uses the name. | Explain how to retrieve the actual cell address of a named range that points to a ListObject and how to trigger formula calculation. | Provide step‑by‑step instructions to rename a ListObject, create a workbook‑level name, and reuse that name in multiple worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

namespace AsposeCellsNamedRangeVLookup
{
    // This C# example demonstrates how to create a workbook, add a ListObject (table) named "ProductsTable", define a workbook‑level named range "ProductRange" that points to the table, insert a VLOOKUP formula using the named range to retrieve the price of "Orange", calculate the formula, output the result, and save the file as an .xlsx document.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a table (header + 3 rows)
                // A1:B1 – headers, A2:B4 – data
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(15);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(8);

                // Add a ListObject (table) that covers the data range A1:B4
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int listObjIdx = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject productsTable = sheet.ListObjects[listObjIdx];
                // Set the table name (DisplayName is the correct property)
                productsTable.DisplayName = "ProductsTable";

                // Create a named range that refers to the table "ProductsTable"
                int nameIdx = workbook.Worksheets.Names.Add("ProductRange");
                Name productRangeName = workbook.Worksheets.Names[nameIdx];
                productRangeName.RefersTo = "=ProductsTable";

                // Retrieve the actual range that the name points to
                Aspose.Cells.Range actualRange = productRangeName.GetRange();
                Console.WriteLine($"Named range 'ProductRange' refers to address: {actualRange.Address}");

                // Use the named range in a VLOOKUP formula.
                // Look up the price of "Orange" using the named range as the table array.
                sheet.Cells["D2"].Formula = "=VLOOKUP(\"Orange\", ProductRange, 2, FALSE)";

                // Calculate formulas so that the VLOOKUP result is evaluated
                workbook.CalculateFormula();

                // Output the result of the VLOOKUP to the console
                Console.WriteLine($"VLOOKUP result for 'Orange': {sheet.Cells["D2"].Value}");

                // Save the workbook (single save)
                string outputPath = "NamedRangeVLookup.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
