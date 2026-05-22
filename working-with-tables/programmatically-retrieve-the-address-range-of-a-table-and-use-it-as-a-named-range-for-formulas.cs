using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTableToNamedRange
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // 3. Populate sample data that will become a table
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue(30);

                // 4. Add a ListObject (Excel Table) covering the data range A1:B4
                //    Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
                Aspose.Cells.Tables.ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable";

                // 5. Retrieve the data range of the table using the DataRange property
                Aspose.Cells.Range tableDataRange = table.DataRange; // e.g., A2:B4 (excluding header)

                // 6. Build the full address string for the named range (including sheet name)
                string namedRangeRef = $"={sheet.Name}!{tableDataRange.Address}";

                // 7. Create a named range that points to the table's data range
                int nameIdx = workbook.Worksheets.Names.Add("SalesData");
                Name namedRange = workbook.Worksheets.Names[nameIdx];
                namedRange.RefersTo = namedRangeRef; // e.g., =DataSheet!A2:B4

                // 8. Use the named range in a formula (sum of Quantity column)
                //    Since the table has two columns, column B holds Quantity.
                //    We'll sum the entire named range; Excel will sum numeric cells only.
                sheet.Cells["D1"].Formula = "=SUM(SalesData)";

                // 9. Calculate formulas to obtain the result
                workbook.CalculateFormula();

                // 10. Output the result to console (optional verification)
                Console.WriteLine($"Named range 'SalesData' refers to: {namedRange.RefersTo}");
                Console.WriteLine($"Sum of SalesData (Quantity): {sheet.Cells["D1"].Value}");

                // 11. Save the workbook (lifecycle rule: save)
                string outputPath = "TableToNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}