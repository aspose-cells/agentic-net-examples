// Title: Aspose.Cells C# – Convert a ListObject (Excel table) to a regular range while preserving formulas
// Description: Demonstrates how to create a workbook, add data and formulas, convert the ListObject (table) back to a normal range using ListObject.ConvertToRange(), and verify that the original formulas remain intact before saving the file.
// Keywords: Aspose.Cells ListObject ConvertToRange | C# convert Excel table to range | preserve formulas Aspose.Cells | Aspose.Cells table to range example | ListObject to range C# | Excel table conversion Aspose | keep formulas after table conversion
// Common Searches: convert ListObject to range Aspose.Cells C# | how to keep formulas when converting Excel table to range | Aspose.Cells ListObject.ConvertToRange usage | remove table but retain formulas Aspose.Cells | C# example converting Excel table back to range
// Developer Intent: Turn an Aspose.Cells ListObject into a normal cell range without losing any existing formulas.
// Use Cases: Replace a table with a plain range to apply custom formatting not supported on ListObjects. | Export workbooks where tables must be removed but calculation logic must stay functional. | Batch‑process worksheets to convert all tables to ranges for compatibility with downstream tools.
// AI Prompts: Generate C# code that converts every ListObject on a worksheet to a regular range while preserving all cell formulas using Aspose.Cells. | Explain the effects of ListObject.ConvertToRange in Aspose.Cells, including which properties are updated during the conversion. | Provide a step‑by‑step guide to verify that formulas remain unchanged after converting an Aspose.Cells table back to a range.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add data and formulas, convert the ListObject (table) back to a normal range using ListObject.ConvertToRange(), and verify that the original formulas remain intact before saving the file.
    public class ListObjectToRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data with formulas
                sheet.Cells["A1"].PutValue("Qty");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["C1"].PutValue("Total");

                sheet.Cells["A2"].PutValue(2);
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].Formula = "=A2*B2";

                sheet.Cells["A3"].PutValue(5);
                sheet.Cells["B3"].PutValue(7);
                sheet.Cells["C3"].Formula = "=A3*B3";

                // Create a ListObject (table) covering the data range
                int tableIndex = sheet.ListObjects.Add("A1", "C3", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Convert the table back to a regular range; formulas are preserved
                table.ConvertToRange();

                // Verify that formulas are still present after conversion
                Console.WriteLine("Formula in C2 after conversion: " + sheet.Cells["C2"].Formula);
                Console.WriteLine("Formula in C3 after conversion: " + sheet.Cells["C3"].Formula);

                // Save the workbook
                workbook.Save("ListObjectToRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ListObjectToRangeDemo.Run();
        }
    }
}
