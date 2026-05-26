using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace DynamicNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data with headers
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue(1);
                cells["B2"].PutValue(10);
                cells["A3"].PutValue(2);
                cells["B3"].PutValue(20);
                cells["A4"].PutValue(3);
                cells["B4"].PutValue(30);

                // Create a table (ListObject) covering the data range (rows 0‑3, cols 0‑1)
                int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "MyTable";

                // Define a dynamic named range that refers to the whole table
                int nameIndex = workbook.Worksheets.Names.Add("MyDynamicRange");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                dynamicName.RefersTo = "=MyTable[#All]";

                // Use the named range in a formula (sum of the Value column)
                cells["D1"].Formula = "=SUM(MyDynamicRange)";

                // Calculate formulas to get the initial sum
                workbook.CalculateFormula();

                // Add a new row to the table
                int newRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;
                sheet.Cells.InsertRows(newRowIndex, 1);
                cells[newRowIndex, 0].PutValue(4); // ID
                cells[newRowIndex, 1].PutValue(40); // Value

                // Expand the table to include the new row (hasHeaders = true)
                table.Resize(
                    table.DataRange.FirstRow,
                    table.DataRange.FirstColumn,
                    table.DataRange.RowCount + 1,
                    table.DataRange.ColumnCount,
                    true);

                // Recalculate formulas after adding the row
                workbook.CalculateFormula();

                // Output the updated sum to console
                Console.WriteLine("Updated sum using dynamic named range: " + cells["D1"].Value);

                // Save the workbook (ensure the directory exists)
                string outputPath = "DynamicNamedRangeDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}