using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDashboardDemo
{
    public class TableHeaderVisibility
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(85);
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(60);

                // Add a ListObject (Excel table) covering the data range
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Apply a built‑in table style (optional, for visual compactness)
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Hide the header row as required for the compact dashboard layout
                table.ShowHeaderRow = false;

                // Define output path
                string outputPath = "DashboardCompactTable.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableHeaderVisibility.Run();
        }
    }
}