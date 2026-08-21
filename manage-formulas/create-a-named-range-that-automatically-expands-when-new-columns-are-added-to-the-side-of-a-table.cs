// Title: Aspose.Cells C# – Create a Named Range That Auto‑Expands When Adding Columns to a Table
// Description: Demonstrates how to build a workbook, define a ListObject (Excel table), create a named range using the structured reference TableName[#All], insert a new column, and verify that the named range automatically includes the added column before saving the file.
// Keywords: Aspose.Cells | C# | .NET | dynamic named range | structured reference | Table[#All] | Excel table | ListObject | auto‑expand range | add column to table | workbook automation
// Common Searches: Aspose.Cells create named range that expands with new columns | C# structured reference Table[#All] auto update | dynamic range for Excel table using Aspose.Cells | how to make a named range follow table schema changes | add column to ListObject and keep named range current
// Developer Intent: Generate a named range that automatically grows to include any columns added to an Excel table.
// Use Cases: Maintain a single reference for all table columns in formulas, charts, or data validations that adapts to schema changes. | Add new data fields to a table without manually updating named ranges in downstream reports. | Export workbooks where external processes rely on a consistent named range that reflects the latest table structure.
// AI Prompts: Write C# code with Aspose.Cells that creates an Excel table, defines a named range using TableName[#All], inserts a new column, and shows the updated range address. | Explain how Table[#All] structured references keep named ranges dynamic in Aspose.Cells when columns are added. | Provide a step‑by‑step tutorial for building a dynamic named range for a ListObject in Aspose.Cells, including verification after column insertion.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicNamedRange
{
    // Demonstrates how to build a workbook, define a ListObject (Excel table), create a named range using the structured reference TableName[#All], insert a new column, and verify that the named range automatically includes the added column before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that will become a table (ListObject)
                // A1:D5 with headers in the first row
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Qty");
                sheet.Cells["D1"].PutValue("Price");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["D2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["D3"].PutValue(0.3);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Cherry");
                sheet.Cells["C4"].PutValue(15);
                sheet.Cells["D4"].PutValue(0.8);

                sheet.Cells["A5"].PutValue(4);
                sheet.Cells["B5"].PutValue("Date");
                sheet.Cells["C5"].PutValue(5);
                sheet.Cells["D5"].PutValue(1.2);

                // Add a ListObject (Excel Table) covering the data range A1:D5
                // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 3, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable"; // Optional: give the table a friendly name

                // Create a named range that refers to the whole table using a structured reference.
                // Structured references like TableName[#All] automatically expand when columns are added.
                int nameIndex = workbook.Worksheets.Names.Add("SalesTableRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                namedRange.RefersTo = "=SalesTable[#All]";

                // Verify the initial address of the named range
                AsposeRange initialRange = namedRange.GetRange();
                Console.WriteLine("Initial named range address: " + initialRange.Address); // Expected: SalesTable[#All]

                // Insert a new column to the right of the table (e.g., after column D)
                // The table will automatically expand to include the new column.
                sheet.Cells.InsertColumn(4); // Inserts before column E (0‑based index)

                // Add a header for the new column and some sample data
                sheet.Cells["E1"].PutValue("Discount");
                sheet.Cells["E2"].PutValue(0.05);
                sheet.Cells["E3"].PutValue(0.10);
                sheet.Cells["E4"].PutValue(0.00);
                sheet.Cells["E5"].PutValue(0.07);

                // After inserting the column, the structured reference should now include column E.
                // Retrieve the range again to see the updated address.
                AsposeRange updatedRange = namedRange.GetRange();
                Console.WriteLine("Updated named range address after column insertion: " + updatedRange.Address);

                // Save the workbook (lifecycle: save)
                string outputPath = "DynamicNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
