// Title: C# – Convert Aspose.Cells ListObject (Table) to a Range while Keeping Formulas
// Description: Demonstrates how to use Aspose.Cells' ConvertToRange method to change a ListObject back to a normal cell range, preserving all formulas and recalculating them before saving the workbook.
// Keywords: Aspose.Cells ConvertToRange | C# ListObject to range | preserve formulas Aspose.Cells | Excel table to range conversion | Aspose.Cells table conversion example
// Common Searches: Aspose.Cells convert table to range C# | keep formulas when converting ListObject | ConvertToRange method example | how to change Excel table back to range using Aspose | C# code to preserve formulas after table conversion
// Developer Intent: Transform a ListObject into a standard cell range without losing any embedded formulas.
// Use Cases: Convert a data table to a range before exporting to ensure formulas remain functional in downstream tools. | Apply custom formatting that tables do not support while retaining calculated values. | Maintain compatibility with older Excel versions that do not recognize tables, preserving all calculations.
// AI Prompts: Write C# code with Aspose.Cells that converts a ListObject to a range and verifies that formulas stay intact. | Explain the steps of the ConvertToRange method and why a workbook.CalculateFormula call is needed afterward. | Create a unit test in C# that asserts formula results are unchanged after converting an Aspose.Cells table to a range.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to use Aspose.Cells' ConvertToRange method to change a ListObject back to a normal cell range, preserving all formulas and recalculating them before saving the workbook.
    public class ListObjectToRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with a header row
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");
                sheet.Cells["D1"].PutValue("Total"); // Column for formula

                // Add some rows of data
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(5);
                sheet.Cells["C3"].PutValue(0.3);

                // Insert a formula that calculates Total = Quantity * Price
                sheet.Cells["D2"].Formula = "=B2*C2";
                sheet.Cells["D3"].Formula = "=B3*C3";

                // Create a ListObject (table) that includes the data and the formula column
                int tableIndex = sheet.ListObjects.Add("A1", "D3", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Optional: set a table style (not required for conversion)
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Convert the ListObject back to a regular range.
                // This operation preserves the existing formulas in the cells.
                table.ConvertToRange();

                // Recalculate formulas to ensure they are evaluated after conversion
                workbook.CalculateFormula();

                // Output the calculated totals to the console for verification
                Console.WriteLine("Total for Apple: " + sheet.Cells["D2"].Value);
                Console.WriteLine("Total for Banana: " + sheet.Cells["D3"].Value);

                // Save the workbook
                string outputPath = "ListObjectConvertedToRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
