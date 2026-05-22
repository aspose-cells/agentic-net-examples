using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsQueryTableConversion
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate sample data ----------
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("John");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Mary");

                // ---------- Create a ListObject (table) ----------
                // In a real scenario this ListObject would be linked to a QueryTable
                int loIndex = sheet.ListObjects.Add("A1", "B3", true);
                ListObject listObject = sheet.ListObjects[loIndex];

                // For demonstration, assume the ListObject has an associated QueryTable
                // (In practice this occurs when the table is created from an external data source)
                QueryTable qt = listObject.QueryTable;

                if (qt != null)
                {
                    // Preserve existing formatting during conversion
                    qt.PreserveFormatting = true;

                    // Capture the current range of the ListObject
                    Aspose.Cells.Range dataRange = listObject.DataRange;
                    int firstRow = dataRange.FirstRow;
                    int firstColumn = dataRange.FirstColumn;
                    int lastRow = firstRow + dataRange.RowCount - 1;
                    int lastColumn = firstColumn + dataRange.ColumnCount - 1;

                    // Convert the ListObject (which contains the QueryTable) to a plain range
                    listObject.ConvertToRange();

                    // Add a new regular ListObject (worksheet table) on the same range
                    int newLoIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
                    ListObject newTable = sheet.ListObjects[newLoIndex];

                    // Example: set a table style
                    newTable.TableStyleType = TableStyleType.TableStyleMedium9;
                }
                else
                {
                    Console.WriteLine("No QueryTable associated with the ListObject. No conversion performed.");
                }

                // ---------- Save the workbook ----------
                string outputPath = "ConvertedQueryTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}