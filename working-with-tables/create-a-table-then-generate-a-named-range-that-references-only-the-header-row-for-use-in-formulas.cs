using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsHeaderNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate header row
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");

                // Populate some data rows
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(2.5);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(1.8);

                // Define the table range (including header)
                int startRow = 0;      // Row index for "A1"
                int startColumn = 0;   // Column index for "A1"
                int endRow = 2;        // Row index for "B3"
                int endColumn = 1;     // Column index for "B3"
                bool hasHeaders = true;

                // Add the ListObject (table) to the worksheet
                int tableIndex = worksheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, hasHeaders);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "ProductTable";

                // Create a named range that refers only to the header row of the table
                // Header row is the first row of the table range (startRow, startColumn) with 1 row and the same column count as the table
                int headerRowCount = 1;
                int headerColumnCount = table.ListColumns.Count; // Number of columns in the table
                Aspose.Cells.Range headerRange = cells.CreateRange(startRow, startColumn, headerRowCount, headerColumnCount);
                headerRange.Name = "ProductHeaders"; // Named range for the header row

                // Example usage of the named range in a formula (count number of header cells)
                cells["C1"].Formula = "=COUNTA(ProductHeaders)";
                workbook.CalculateFormula();
                Console.WriteLine("Header count (should be 2): " + cells["C1"].IntValue);

                // Save the workbook (ensure the directory exists)
                string outputPath = "HeaderNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}