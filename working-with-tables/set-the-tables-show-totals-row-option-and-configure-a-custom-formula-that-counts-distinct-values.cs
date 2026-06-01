using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalsRowDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data with a header and some duplicate values
                cells["A1"].PutValue("Category");   // Header
                cells["A2"].PutValue("Apple");
                cells["A3"].PutValue("Banana");
                cells["A4"].PutValue("Apple");
                cells["A5"].PutValue("Orange");
                cells["A6"].PutValue("Banana");

                // Add a table that includes the data range (A1:A6) and has a header row
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 5, 0, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Enable the totals row for the table
                table.ShowTotals = true;

                // Configure the first column (Category) to use a custom totals calculation
                ListColumn categoryColumn = table.ListColumns[0];
                categoryColumn.TotalsCalculation = TotalsCalculation.Custom;

                // Set a custom formula that counts distinct values in the column.
                // The formula uses COUNTIF to count each occurrence and then sums the reciprocals.
                // This is an array‑style formula; Aspose.Cells accepts it as a regular formula string.
                string distinctCountFormula = "=SUM(1/COUNTIF([Category],[Category]))";
                // isR1C1 = false (A1 style), isLocal = false (invariant)
                categoryColumn.SetCustomTotalsRowFormula(distinctCountFormula, false, false);

                // Optionally set a label for the totals row in the second column (if it exists)
                // Here we add a second column just to demonstrate the label.
                cells["B1"].PutValue("Value");
                cells["B2"].PutValue(10);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(30);
                cells["B5"].PutValue(40);
                cells["B6"].PutValue(50);

                // Expand the table to include the second column (hasHeaders = true)
                table.Resize(0, 0, 5, 1, true);

                // Set a label for the totals row of the second column
                table.ListColumns[1].TotalsRowLabel = "Distinct Count";

                // Save the workbook
                string outputPath = "TableWithDistinctCountTotals.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}