// Title: Aspose.Cells C# – Create a ListObject table and a header‑only named range
// Description: This example shows how to build a workbook, add a ListObject (table) with a three‑column header, define a named range that points exclusively to the header row, use the range in a COUNTA formula, and save the file. The code demonstrates creating the range programmatically and referencing it in formulas.
// Keywords: Aspose.Cells C# table header named range | Create ListObject Aspose.Cells | header‑only named range .NET | Aspose.Cells COUNTA formula example | C# workbook named range for table headers | Aspose.Cells table header range
// Common Searches: Aspose.Cells create named range for table header row | C# ListObject header range Aspose.Cells | How to reference only the header of a table in Aspose.Cells | Aspose.Cells named range COUNTA header example | Define table header range in .NET workbook
// Developer Intent: Generate a worksheet table and a named range that references only its header row for formula use.
// Use Cases: Count header columns with a COUNTA formula referencing the named range. | Supply the header range to data‑validation lists or drop‑downs on other sheets. | Use the header named range in VLOOKUP, MATCH, or INDEX formulas to locate columns dynamically.
// AI Prompts: Write C# code with Aspose.Cells that creates a ListObject table and a named range for its header row, then applies a COUNTA formula to count the headers. | Explain how to extend the header‑only named range to cover multiple header rows in an Aspose.Cells workbook. | Show how to retrieve the header named range programmatically and use it in a VLOOKUP formula across worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This example shows how to build a workbook, add a ListObject (table) with a three‑column header, define a named range that points exclusively to the header row, use the range in a COUNTA formula, and save the file. The code demonstrates creating the range programmatically and referencing it in formulas.
    public class TableHeaderNamedRangeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header row (A1:C1) and some data rows
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue("Fruit");
            cells["C2"].PutValue(1.2);

            cells["A3"].PutValue("Carrot");
            cells["B3"].PutValue("Vegetable");
            cells["C3"].PutValue(0.8);

            // Define the table range (including header row)
            int startRow = 0;      // Row 0 -> A1
            int startColumn = 0;   // Column 0 -> A
            int endRow = 2;        // Row 2 -> third row (A3:C3)
            int endColumn = 2;     // Column 2 -> C

            // Add a ListObject (table) that has headers
            int tableIndex = worksheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "ProductTable";

            // Calculate the number of columns in the table
            int columnCount = endColumn - startColumn + 1;

            // Create a named range that refers only to the header row of the table
            AsposeRange headerRange = cells.CreateRange(startRow, startColumn, 1, columnCount);
            headerRange.Name = "ProductTableHeaders";

            // Example usage of the named range in a formula (count headers)
            cells["E1"].Formula = "=COUNTA(ProductTableHeaders)";
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "TableHeaderNamedRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
