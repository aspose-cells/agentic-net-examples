using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableHeaderFix
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

                // Populate sample data without a header row
                // Table will occupy A1:C3 (3 rows, 3 columns) with no header row
                worksheet.Cells["A1"].PutValue("Apple");
                worksheet.Cells["B1"].PutValue(10);
                worksheet.Cells["C1"].PutValue(1.5);
                worksheet.Cells["A2"].PutValue("Banana");
                worksheet.Cells["B2"].PutValue(20);
                worksheet.Cells["C2"].PutValue(2.0);
                worksheet.Cells["A3"].PutValue("Cherry");
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["C3"].PutValue(2.5);

                // Add a ListObject (table) without headers (hasHeaders = false)
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 2, false);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "FruitTable";

                // Iterate through all tables in the worksheet
                foreach (ListObject lo in worksheet.ListObjects)
                {
                    // Insert a new row at the start of the table to become the header row
                    worksheet.Cells.InsertRows(lo.StartRow, 1);

                    // Fill the new header row with generic column names (Column1, Column2, ...)
                    int columnCount = lo.ListColumns.Count;
                    for (int col = 0; col < columnCount; col++)
                    {
                        worksheet.Cells[lo.StartRow, lo.StartColumn + col].PutValue($"Column{col + 1}");
                    }

                    // Update the ListObject's column names to match the new header cells
                    lo.UpdateColumnName();
                }

                // Save the workbook
                workbook.Save("TableWithHeaders.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}