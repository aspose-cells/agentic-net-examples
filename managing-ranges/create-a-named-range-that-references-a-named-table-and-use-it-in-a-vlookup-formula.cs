using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsNamedRangeVLookup
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data that will become a table
                // Header row
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");

                // Data rows
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(30);

                // Convert the range A1:B4 into a table (ListObject)
                // The table will be named "ProductsTable"
                int firstRow = 0;   // zero‑based index for row 1
                int firstCol = 0;   // zero‑based index for column A
                int totalRows = 4;  // includes header
                int totalCols = 2;

                // Add returns the index of the created ListObject
                int tableIdx = sheet.ListObjects.Add(
                    firstRow,
                    firstCol,
                    firstRow + totalRows - 1,
                    firstCol + totalCols - 1,
                    true);

                // Retrieve the ListObject instance
                ListObject table = sheet.ListObjects[tableIdx];
                // Set the display name of the table (used in formulas)
                table.DisplayName = "ProductsTable";
                // Ensure the header row is visible
                table.ShowHeaderRow = true;

                // Create a named range that references the table
                // The named range "ProductsRange" will point to the entire table
                int nameIdx = workbook.Worksheets.Names.Add("ProductsRange");
                Name namedRange = workbook.Worksheets.Names[nameIdx];
                // RefersTo must start with '='
                namedRange.RefersTo = "=ProductsTable";

                // Use the named range in a VLOOKUP formula
                // Lookup "Banana" and return its price (second column)
                cells["D2"].Formula = "=VLOOKUP(\"Banana\", ProductsRange, 2, FALSE)";

                // Calculate formulas so that the VLOOKUP result is evaluated
                workbook.CalculateFormula();

                // Output the result to console (optional verification)
                Console.WriteLine("VLOOKUP result for 'Banana': " + cells["D2"].Value);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "NamedRangeVLookupDemo.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}