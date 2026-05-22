using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableExample
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

                // Populate header and a couple of data rows to form a table
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Price");

                cells["A2"].PutValue(1);
                cells["B2"].PutValue("Laptop");
                cells["C2"].PutValue(999.99m);

                cells["A3"].PutValue(2);
                cells["B3"].PutValue("Monitor");
                cells["C3"].PutValue(249.99m);

                // Create a ListObject (Excel table) covering the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "C3", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Dictionary containing values for the new row
                var newRowData = new Dictionary<string, object>
                {
                    { "ID", 3 },
                    { "Product", "Keyboard" },
                    { "Price", 49.99m }
                };

                // Determine the row offset for the new row (append at the end)
                // DataRange includes the header row, so subtract 1 to get existing data rows count
                int newRowOffset = table.DataRange.RowCount - 1;

                // Populate each column using the dictionary values
                foreach (KeyValuePair<string, object> kvp in newRowData)
                {
                    // Find the column index by header name
                    int columnIndex = -1;
                    for (int i = 0; i < table.ListColumns.Count; i++)
                    {
                        if (string.Equals(table.ListColumns[i].Name, kvp.Key, StringComparison.OrdinalIgnoreCase))
                        {
                            columnIndex = i;
                            break;
                        }
                    }

                    // If the column exists, put the value; otherwise ignore
                    if (columnIndex >= 0)
                    {
                        table.PutCellValue(newRowOffset, columnIndex, kvp.Value);
                    }
                }

                // Save the workbook
                string outputPath = "TableWithNewRow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}